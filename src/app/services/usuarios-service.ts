import { Injectable } from "@angular/core";
import { HttpClient, HttpHeaders } from "@angular/common/http";
import { Observable, from } from "rxjs";
import { firstValueFrom } from "rxjs";
import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";
import { environment } from "../../environments/environment";

/** Usuario tal como se usa en la UI de administración */
export interface Usuario {
  id: string;
  nombre: string;
  correo: string;
  rol: string;
  estado: string;
}

/** Payload para crear un usuario en Entra ID */
export interface CrearUsuarioPayload {
  nombre: string;
  correo: string;
  rol: string;
  passwordTemporal: string;
}

/** Payload para actualizar nombre y/o rol */
export interface ActualizarUsuarioPayload {
  nombre?: string;
  rol?: string;
}

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Permisos Graph (delegados) que el admin debe conceder en Azure Portal
 * (App registration del FRONTEND → API permissions → Grant admin consent):
 * - User.ReadWrite.All
 * - Application.Read.All
 * - AppRoleAssignment.ReadWrite.All
 */
export const GRAPH_USER_SCOPES = [
  "User.ReadWrite.All",
  "Application.Read.All",
  "AppRoleAssignment.ReadWrite.All",
];

const ROLES_APP = [
  "ROLE_ADMINISTRADOR",
  "ROLE_OPERADOR",
  "ROLE_CLIENTE",
] as const;

const DEFAULT_APPROLE_ID = "00000000-0000-0000-0000-000000000000";

interface AppRoleInfo {
  id: string;
  value: string;
  displayName?: string;
}

interface ServicePrincipalInfo {
  id: string;
  appId: string;
  appRoles: AppRoleInfo[];
}

@Injectable({
  providedIn: "root",
})
export class UsuariosService {
  private spCache = new Map<string, ServicePrincipalInfo>();

  constructor(
    private http: HttpClient,
    private auth: AuthService,
    private msal: MsalService,
  ) {}

  private get apiAppId(): string {
    const fromEnv = (environment as any).apiConfig?.apiAppId as
      | string
      | undefined;
    if (fromEnv) {
      return fromEnv;
    }
    const scope = environment.apiConfig.scopes?.[0] || "";
    const match = scope.match(/api:\/\/([^/]+)/);
    return match ? match[1] : "";
  }

  private get spaAppId(): string {
    return environment.msalConfig?.auth?.clientId || "";
  }

  private async getGraphToken(): Promise<string | null> {
    const account = this.auth.getActiveAccount();
    if (!account) {
      return null;
    }

    try {
      const result = await this.msal.instance.acquireTokenSilent({
        scopes: GRAPH_USER_SCOPES,
        account,
      });
      return result.accessToken;
    } catch (error) {
      console.warn(
        "No se pudo obtener token de Graph en silencio. Se intentará interactivo.",
        error,
      );
      try {
        const result = await this.msal.instance.acquireTokenPopup({
          scopes: GRAPH_USER_SCOPES,
          account,
        });
        return result.accessToken;
      } catch (popupError) {
        console.error("Error obteniendo token de Graph:", popupError);
        return null;
      }
    }
  }

  private async authHeaders(): Promise<HttpHeaders> {
    const token = await this.getGraphToken();
    if (!token) {
      throw new Error(
        "No se pudo obtener token de Microsoft Graph. En Azure Portal, en el App registration del FRONTEND, agrega los permisos delegados User.ReadWrite.All, Application.Read.All y AppRoleAssignment.ReadWrite.All y pulsa Grant admin consent.",
      );
    }
    return new HttpHeaders({
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    });
  }

  private async graphGet<T>(url: string, headers: HttpHeaders): Promise<T> {
    return firstValueFrom(this.http.get<T>(url, { headers }));
  }

  private async graphPost<T>(
    url: string,
    body: unknown,
    headers: HttpHeaders,
  ): Promise<T> {
    return firstValueFrom(this.http.post<T>(url, body, { headers }));
  }

  private async graphPatch(
    url: string,
    body: unknown,
    headers: HttpHeaders,
  ): Promise<void> {
    await firstValueFrom(this.http.patch(url, body, { headers }));
  }

  private async graphDelete(url: string, headers: HttpHeaders): Promise<void> {
    await firstValueFrom(this.http.delete(url, { headers }));
  }

  private async getServicePrincipalByAppId(
    appId: string,
    headers: HttpHeaders,
  ): Promise<ServicePrincipalInfo | null> {
    if (!appId) {
      return null;
    }
    if (this.spCache.has(appId)) {
      return this.spCache.get(appId)!;
    }

    const res = await this.graphGet<{ value: ServicePrincipalInfo[] }>(
      `${GRAPH_BASE}/servicePrincipals?$filter=appId eq '${appId}'&$select=id,appId,appRoles`,
      headers,
    );

    const sp = res.value?.[0];
    if (!sp) {
      return null;
    }

    const info: ServicePrincipalInfo = {
      id: sp.id,
      appId: sp.appId || appId,
      appRoles: (sp.appRoles || [])
        .filter((r) => r.id && r.id !== DEFAULT_APPROLE_ID)
        .map((r) => ({
          id: r.id,
          value: r.value,
          displayName: r.displayName,
        })),
    };

    this.spCache.set(appId, info);
    return info;
  }

  private encontrarAppRole(
    sp: ServicePrincipalInfo,
    rolDeseado: string,
  ): AppRoleInfo | null {
    const target = this.normalizarRol(rolDeseado);

    const exact = sp.appRoles.find(
      (r) => (r.value || "").toUpperCase() === target,
    );
    if (exact) {
      return exact;
    }

    const porNombre = sp.appRoles.find((r) => {
      const v = `${r.value || ""} ${r.displayName || ""}`.toUpperCase();
      if (target === "ROLE_ADMINISTRADOR") {
        return v.includes("ADMIN");
      }
      if (target === "ROLE_OPERADOR") {
        return v.includes("OPERADOR") || v.includes("OPERATOR");
      }
      if (target === "ROLE_CLIENTE") {
        return v.includes("CLIENTE") || v.includes("CLIENT");
      }
      return false;
    });

    return porNombre || null;
  }

  /**
   * Asigna el App Role al usuario en un Enterprise App y quita los otros
   * roles de ESA misma app. El claim "roles" del JWT se actualiza en el
   * próximo login (sesión cerrada + entrar de nuevo).
   */
  private async asignarRolEnSp(
    userId: string,
    rolDeseado: string,
    sp: ServicePrincipalInfo,
    headers: HttpHeaders,
  ): Promise<boolean> {
    const appRole = this.encontrarAppRole(sp, rolDeseado);
    if (!appRole) {
      const disponibles = sp.appRoles
        .map((r) => r.value || r.displayName)
        .filter(Boolean)
        .join(", ");
      console.warn(
        `La app ${sp.appId} no tiene el App Role "${rolDeseado}". Definidos: ${disponibles || "(ninguno)"}`,
      );
      return false;
    }

    const assignments = await this.graphGet<{
      value: Array<{
        id: string;
        appRoleId: string;
        resourceId: string;
        principalId: string;
      }>;
    }>(
      `${GRAPH_BASE}/users/${userId}/appRoleAssignments?$top=100`,
      headers,
    );

    const deEstaApi = (assignments.value || []).filter(
      (a) => a.resourceId === sp.id,
    );

    const yaTieneElRol = deEstaApi.some((a) => a.appRoleId === appRole.id);

    for (const a of deEstaApi) {
      if (a.appRoleId !== appRole.id) {
        await this.graphDelete(
          `${GRAPH_BASE}/users/${userId}/appRoleAssignments/${a.id}`,
          headers,
        );
      }
    }

    if (!yaTieneElRol) {
      await this.graphPost(
        `${GRAPH_BASE}/users/${userId}/appRoleAssignments`,
        {
          principalId: userId,
          resourceId: sp.id,
          appRoleId: appRole.id,
        },
        headers,
      );
    }

    return true;
  }

  /**
   * Intenta asignar el rol en la API y también en el frontend (SPA),
   * porque el JWT de la API usa los App Roles de la API.
   */
  private async asignarRolApp(
    userId: string,
    rolDeseado: string,
    headers: HttpHeaders,
  ): Promise<void> {
    const appIds = [this.apiAppId, this.spaAppId].filter(
      (id, idx, arr) => !!id && arr.indexOf(id) === idx,
    );

    const errores: string[] = [];
    let asignado = false;

    for (const appId of appIds) {
      try {
        const sp = await this.getServicePrincipalByAppId(appId, headers);
        if (!sp) {
          errores.push(`No existe Enterprise Application para ${appId}`);
          continue;
        }
        if (!sp.appRoles.length) {
          errores.push(`La app ${appId} no tiene App Roles definidos`);
          continue;
        }
        const ok = await this.asignarRolEnSp(userId, rolDeseado, sp, headers);
        if (ok) {
          asignado = true;
        } else {
          errores.push(
            `No se encontró el App Role "${rolDeseado}" en ${appId}`,
          );
        }
      } catch (err: any) {
        const msg =
          err?.error?.error?.message ||
          err?.message ||
          `Error asignando rol en ${appId}`;
        errores.push(msg);
        console.error(err);
      }
    }

    if (!asignado) {
      throw new Error(
        `No se pudo asignar el App Role "${rolDeseado}" (el JWT no va a cambiar). ` +
          errores.join(" | ") +
          " Crea los App Roles ROLE_ADMINISTRADOR, ROLE_OPERADOR y ROLE_CLIENTE " +
          "en el App registration de la API (4ea351c0-ef77-41aa-acac-dd63f5d0648f), " +
          "con Allowed member types = Users/Groups, y concede AppRoleAssignment.ReadWrite.All.",
      );
    }
  }

  private async mapaRolesPorUsuario(
    headers: HttpHeaders,
  ): Promise<Map<string, string>> {
    const mapa = new Map<string, string>();
    const appIds = [this.apiAppId, this.spaAppId].filter(
      (id, idx, arr) => !!id && arr.indexOf(id) === idx,
    );

    for (const appId of appIds) {
      const sp = await this.getServicePrincipalByAppId(appId, headers);
      if (!sp) {
        continue;
      }

      const res = await this.graphGet<{
        value: Array<{
          principalId: string;
          appRoleId: string;
          principalType?: string;
        }>;
      }>(
        `${GRAPH_BASE}/servicePrincipals/${sp.id}/appRoleAssignedTo?$top=999`,
        headers,
      );

      const rolPorId = new Map<string, string>();
      for (const r of sp.appRoles) {
        rolPorId.set(r.id, this.normalizarRol(r.value || r.displayName || ""));
      }

      for (const a of res.value || []) {
        if (a.principalType && a.principalType !== "User") {
          continue;
        }
        const rol = rolPorId.get(a.appRoleId);
        if (rol && ROLES_APP.includes(rol as (typeof ROLES_APP)[number])) {
          // Prioriza la API: si ya hay un rol de la API, no lo pisa el SPA
          if (appId === this.apiAppId || !mapa.has(a.principalId)) {
            mapa.set(a.principalId, rol);
          }
        }
      }
    }

    return mapa;
  }

  listarUsuarios(): Observable<Usuario[]> {
    return from(this.listarUsuariosAsync());
  }

  private async listarUsuariosAsync(): Promise<Usuario[]> {
    const headers = await this.authHeaders();
    const res = await this.graphGet<{ value: any[] }>(
      `${GRAPH_BASE}/users?$select=id,displayName,mail,userPrincipalName,accountEnabled,jobTitle&$top=100`,
      headers,
    );

    let rolesMap = new Map<string, string>();
    try {
      rolesMap = await this.mapaRolesPorUsuario(headers);
    } catch (err) {
      console.warn(
        "No se pudieron leer App Role assignments; se usará jobTitle como respaldo.",
        err,
      );
    }

    return (res.value || []).map((u: any) => {
      const rolAsignado = rolesMap.get(u.id);
      return this.mapGraphUser(u, rolAsignado);
    });
  }

  crearUsuario(payload: CrearUsuarioPayload): Observable<Usuario | null> {
    return from(this.crearUsuarioAsync(payload));
  }

  private async crearUsuarioAsync(
    payload: CrearUsuarioPayload,
  ): Promise<Usuario | null> {
    const headers = await this.authHeaders();

    const creado = await this.graphPost<any>(
      `${GRAPH_BASE}/users`,
      {
        accountEnabled: true,
        displayName: payload.nombre,
        mailNickname: payload.correo.split("@")[0],
        userPrincipalName: payload.correo,
        jobTitle: payload.rol,
        passwordProfile: {
          forceChangePasswordNextSignIn: true,
          password: payload.passwordTemporal,
        },
      },
      headers,
    );

    await this.asignarRolApp(creado.id, payload.rol, headers);

    return this.mapGraphUser(creado, payload.rol);
  }

  actualizarUsuario(
    id: string,
    payload: ActualizarUsuarioPayload,
  ): Observable<Usuario | null> {
    return from(this.actualizarUsuarioAsync(id, payload));
  }

  private async actualizarUsuarioAsync(
    id: string,
    payload: ActualizarUsuarioPayload,
  ): Promise<Usuario | null> {
    const headers = await this.authHeaders();

    const body: Record<string, string> = {};
    if (payload.nombre !== undefined) {
      body["displayName"] = payload.nombre;
    }
    if (payload.rol !== undefined) {
      body["jobTitle"] = payload.rol;
    }

    if (Object.keys(body).length > 0) {
      await this.graphPatch(`${GRAPH_BASE}/users/${id}`, body, headers);
    }

    if (payload.rol !== undefined) {
      await this.asignarRolApp(id, payload.rol, headers);
    }

    const u = await this.graphGet<any>(
      `${GRAPH_BASE}/users/${id}?$select=id,displayName,mail,userPrincipalName,accountEnabled,jobTitle`,
      headers,
    );

    return this.mapGraphUser(u, payload.rol);
  }

  private mapGraphUser(u: any, rolApp?: string): Usuario {
    const rol = rolApp || u.jobTitle || "ROLE_CLIENTE";
    return {
      id: u.id,
      nombre: u.displayName || "",
      correo: u.mail || u.userPrincipalName || "",
      rol: this.normalizarRol(rol),
      estado: u.accountEnabled === false ? "INACTIVO" : "ACTIVO",
    };
  }

  private normalizarRol(rol: string): string {
    const r = (rol || "").toUpperCase().trim();
    if (
      r === "ROLE_ADMINISTRADOR" ||
      r === "ROLE_OPERADOR" ||
      r === "ROLE_CLIENTE"
    ) {
      return r;
    }
    if (r.includes("ADMIN")) return "ROLE_ADMINISTRADOR";
    if (r.includes("OPERADOR") || r.includes("OPERATOR")) return "ROLE_OPERADOR";
    if (r.includes("CLIENTE") || r.includes("CLIENT")) return "ROLE_CLIENTE";
    return "ROLE_CLIENTE";
  }
}