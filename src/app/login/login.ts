import { Component, Inject, OnInit } from "@angular/core";
import { CommonModule } from "@angular/common";
import { Router } from "@angular/router";

import {
  MsalService,
  MsalBroadcastService,
  MSAL_GUARD_CONFIG,
  MsalGuardConfiguration,
} from "@azure/msal-angular";

import { RedirectRequest, InteractionStatus } from "@azure/msal-browser";

import { filter } from "rxjs/operators";

@Component({
  selector: "app-login",
  standalone: true,
  imports: [CommonModule],
  template: `
    <div style="padding:40px; text-align:center; font-family:Arial">
      <h1>Pedidos360</h1>

      <p *ngIf="!isLoggedIn">Inicia sesión para continuar</p>

      <p *ngIf="isLoggedIn">Usuario autenticado. Redirigiendo...</p>

      <button
        *ngIf="!isLoggedIn"
        (click)="login()"
        style="
          padding:12px 24px;
          background:#0078d4;
          color:white;
          border:none;
          border-radius:4px;
          cursor:pointer;
        "
      >
        Iniciar sesión con Microsoft
      </button>
    </div>
  `,
})
export class LoginComponent implements OnInit {
  isLoggedIn = false;

  constructor(
    @Inject(MSAL_GUARD_CONFIG)
    private msalGuardConfig: MsalGuardConfiguration,

    private authService: MsalService,
    private msalBroadcastService: MsalBroadcastService,
    private router: Router,
  ) {}

  ngOnInit(): void {
    console.log("LOGIN COMPONENT: ngOnInit ejecutado");

    this.msalBroadcastService.inProgress$
      .pipe(
        filter(
          (status: InteractionStatus) => status === InteractionStatus.None,
        ),
      )
      .subscribe(() => {
        console.log("MSAL terminó de procesar la autenticación");

        const accounts = this.authService.instance.getAllAccounts();

        console.log("Cuentas encontradas:", accounts);

        if (accounts.length > 0) {
          this.isLoggedIn = true;

          this.authService.instance.setActiveAccount(accounts[0]);

          const tokenRequest = {
            scopes: [
              "api://4ea351c0-ef77-41aa-acac-dd63f5d0648f/access_as_user",
            ],
            account: accounts[0],
          };

          this.authService.acquireTokenSilent(tokenRequest).subscribe({
            next: (result) => {
              console.log("✅ Access token del BFF obtenido");

              console.log("Scopes concedidos:", result.scopes);

              const payload = JSON.parse(
                atob(result.accessToken.split(".")[1]),
              );

              console.log("Claims del ACCESS TOKEN del BFF:", payload);

              console.log("Roles del BFF:", payload.roles);

              console.log("Audiencia del BFF:", payload.aud);
            },

            error: (error) => {
              console.error("❌ Error obteniendo token del BFF:", error);
            },
          });

          console.log("Usuario autenticado:", accounts[0].username);

          console.log("Claims de la cuenta:", accounts[0].idTokenClaims);

          setTimeout(() => {
            this.router.navigate(["/dashboard"]);
          }, 1000);
        }
      });
  }

  login(): void {
    console.log("Iniciando login con MSAL");

    if (this.msalGuardConfig.authRequest) {
      this.authService.loginRedirect({
        ...this.msalGuardConfig.authRequest,
      } as RedirectRequest);
    } else {
      this.authService.loginRedirect();
    }
  }
}
