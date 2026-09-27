import { Injectable } from "@angular/core";

import { MsalService } from "@azure/msal-angular";

import { AccountInfo, SilentRequest } from "@azure/msal-browser";

import { environment } from "../../environments/environment";

@Injectable({
  providedIn: "root",
})
export class AuthService {
  private roles: string[] = [];
  private rolesLoaded = false;

  constructor(private msalService: MsalService) {}

  // CUENTA ACTIVA
  getActiveAccount(): AccountInfo | null {
    let account = this.msalService.instance.getActiveAccount();

    if (!account) {
      const accounts = this.msalService.instance.getAllAccounts();

      if (accounts.length > 0) {
        this.msalService.instance.setActiveAccount(accounts[0]);

        account = accounts[0];
      }
    }

    return account;
  }

  // SESIÓN
  isLoggedIn(): boolean {
    return !!this.getActiveAccount();
  }

  /**
   * Access token de la API.
   * forceRefresh=true pide un token nuevo a Azure para que traiga los App Roles
   * actualizados (si no, MSAL reutiliza un JWT viejo de hasta ~1 hora).
   */
  async getAccessToken(forceRefresh = false): Promise<string | null> {
    const account = this.getActiveAccount();

    if (!account) {
      return null;
    }

    const request: SilentRequest = {
      scopes: environment.apiConfig.scopes,
      account: account,
      forceRefresh,
    };

    try {
      const result =
        await this.msalService.instance.acquireTokenSilent(request);

      return result.accessToken;
    } catch (error) {
      console.error("Error obteniendo token del BFF:", error);

      return null;
    }
  }

  // ROLES
  async getRoles(): Promise<string[]> {
    // Siempre re-leer el token: si un admin cambió el App Role, el JWT cacheado
    // seguiría mostrando el rol anterior.
    this.rolesLoaded = false;
    this.roles = [];

    const account = this.getActiveAccount();
    const collected = new Set<string>();

    const accessToken = await this.getAccessToken(true);

    if (accessToken) {
      this.extraerRolesDeJwt(accessToken).forEach((r) => collected.add(r));
    }

    const idClaims = account?.idTokenClaims as
      | Record<string, unknown>
      | undefined;
    this.extraerRolesDeClaims(idClaims).forEach((r) => collected.add(r));

    this.roles = Array.from(collected);
    this.rolesLoaded = true;

    console.log("Roles leídos del token:", this.roles);

    return this.roles;
  }

  private extraerRolesDeJwt(token: string): string[] {
    try {
      const payload = JSON.parse(atob(token.split(".")[1]));
      return this.extraerRolesDeClaims(payload);
    } catch (error) {
      console.error("Error leyendo roles del token:", error);
      return [];
    }
  }

  private extraerRolesDeClaims(
    claims: Record<string, unknown> | undefined | null,
  ): string[] {
    if (!claims) {
      return [];
    }

    const raw = claims["roles"] ?? claims["role"];

    if (Array.isArray(raw)) {
      return raw.filter((r) => typeof r === "string") as string[];
    }

    if (typeof raw === "string") {
      return [raw];
    }

    return [];
  }

  // COMPROBAR ROLES
  async hasRole(role: string): Promise<boolean> {
    const roles = await this.getRoles();

    return roles.includes(role);
  }

  async hasAnyRole(roles: string[]): Promise<boolean> {
    const userRoles = await this.getRoles();

    return roles.some((role) => userRoles.includes(role));
  }

  // LIMPIAR DATOS (llamar al hacer login/logout)
  clearAuthState(): void {
    this.roles = [];
    this.rolesLoaded = false;
  }
}