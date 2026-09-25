import { Component, OnInit } from "@angular/core";
import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";
import { ChangeDetectorRef } from '@angular/core';

@Component({
  selector: "app-dashboard",
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <div style="padding:30px; font-family:Arial">
      <h2>Dashboard - Pedidos360</h2>

      <p>Usuario: {{ nombre }}</p>

      <p>
        Roles:
        {{ roles.join(", ") || "Sin roles" }}
      </p>

      <div style="margin:20px 0">
        <a routerLink="/orders" style="margin-right:15px"> Pedidos </a>

        <a
          *ngIf="esAdmin || esOperador"
          routerLink="/catalog"
          style="margin-right:15px"
        >
          Catálogo
        </a>

        <button (click)="logout()">Cerrar sesión</button>
      </div>

      <div
        *ngIf="esAdmin"
        style="
          background:#e8f5e9;
          padding:15px;
          margin-top:20px
        "
      >
        <h3>Vista Administrador</h3>

        <p>Puedes ver todo el sistema, pedidos y catálogo.</p>
      </div>

      <div
        *ngIf="esOperador"
        style="
          background:#fff3e0;
          padding:15px;
          margin-top:20px
        "
      >
        <h3>Vista Operador</h3>

        <p>Gestionas pedidos y estados. También puedes ver/editar catálogo.</p>
      </div>

      <div
        *ngIf="esCliente"
        style="
          background:#e3f2fd;
          padding:15px;
          margin-top:20px
        "
      >
        <h3>Vista Cliente</h3>

        <p>Puedes crear y ver tus propios pedidos. Solo consulta catálogo.</p>
      </div>
    </div>
  `,
})
export class DashboardComponent implements OnInit {
  nombre = "";

  roles: string[] = [];

  esAdmin = false;
  esOperador = false;
  esCliente = false;

  constructor(
  private auth: AuthService,
  private msal: MsalService,
  private cdr: ChangeDetectorRef
) {}

  async ngOnInit() {

  const account =
    this.auth.getActiveAccount();

  this.nombre =
    account?.name ||
    account?.username ||
    '';

  const roles =
    await this.auth.getRoles();

  this.roles = roles;

  this.esAdmin =
    roles.includes('ROLE_ADMINISTRADOR');

  this.esOperador =
    roles.includes('ROLE_OPERADOR');

  this.esCliente =
    roles.includes('ROLE_CLIENTE');

  console.log('Roles del Dashboard:', this.roles);
  console.log('Admin:', this.esAdmin);
  console.log('Operador:', this.esOperador);
  console.log('Cliente:', this.esCliente);

  this.cdr.detectChanges();
}

  logout() {
    this.msal.logoutRedirect();
  }
}
