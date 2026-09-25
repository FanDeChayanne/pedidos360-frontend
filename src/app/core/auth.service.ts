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

  // ACCESS TOKEN DEL BFF
  async getAccessToken(): Promise<string | null> {
    const account = this.getActiveAccount();

    if (!account) {
      return null;
    }

    const request: SilentRequest = {
      scopes: environment.apiConfig.scopes,
      account: account,
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
    if (this.rolesLoaded) {
      return this.roles;
    }

    const accessToken = await this.getAccessToken();

    if (!accessToken) {
      return [];
    }

    try {
      const payload = JSON.parse(atob(accessToken.split(".")[1]));

      const tokenRoles = payload.roles;

      if (Array.isArray(tokenRoles)) {
        this.roles = tokenRoles;
      } else if (typeof tokenRoles === "string") {
        this.roles = [tokenRoles];
      } else {
        this.roles = [];
      }

      this.rolesLoaded = true;

      return this.roles;
    } catch (error) {
      console.error("Error leyendo roles del token:", error);

      return [];
    }
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

  // LIMPIAR DATOS
  clearAuthState(): void {
    this.roles = [];
    this.rolesLoaded = false;
  }
}
