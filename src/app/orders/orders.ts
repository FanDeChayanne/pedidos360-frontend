import { Component, OnInit, ChangeDetectorRef } from "@angular/core";

import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";

@Component({
  selector: "app-orders",
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <!-- ========== NAVBAR (igual que Dashboard) ========== -->
    <nav class="top-navbar">
      <div class="nav-container">
        <div class="nav-left">
          <a class="nav-item" routerLink="/dashboard">
            <i class="bi bi-house-door-fill"></i>
            Dashboard
          </a>

          <a class="nav-item active" routerLink="/orders">
            <i class="bi bi-cart3"></i>
            Pedidos
          </a>

          <a
            class="nav-item"
            *ngIf="esAdmin || esOperador"
            routerLink="/catalog"
          >
            <i class="bi bi-journal-text"></i>
            Catálogo
          </a>

          <!-- Usuarios (solo Admin) -->
          <a class="nav-item" *ngIf="esAdmin" routerLink="/usuarios">
            <i class="bi bi-people"></i>
            Usuarios
          </a>

          <a class="nav-item" href="javascript:void(0)">
            <i class="bi bi-box"></i>
            TEST
          </a>
        </div>

        <div class="nav-right">
          <div class="user-info">
            <div class="avatar">
              {{ nombre ? nombre.charAt(0).toUpperCase() : "U" }}
            </div>
            <div class="user-text">
              <div class="user-name">{{ nombre || "Usuario" }}</div>
              <div class="user-role">{{ roles.join(", ") || "Sin roles" }}</div>
            </div>
          </div>

          <button class="btn-logout" (click)="logout()">
            <i class="bi bi-box-arrow-right"></i>
            Cerrar sesión
          </button>
        </div>
      </div>
    </nav>

    <!-- ========== CONTENIDO ORIGINAL ========== -->
    <main class="main-content">
      <div class="p-4">
        <h2 class="mb-3">Pedidos</h2>

        <div *ngIf="esCliente" class="card border-0 shadow-sm mb-4">
          <div class="card-body">
            <h5 class="card-title">Mis pedidos (Cliente)</h5>
            <p class="card-text">
              Aquí el cliente ve solo sus pedidos y puede crear uno nuevo.
            </p>
            <button class="btn btn-primary btn-sm">Crear pedido (mock)</button>
          </div>
        </div>

        <div *ngIf="esOperador || esAdmin" class="card border-0 shadow-sm mb-4">
          <div class="card-body">
            <h5 class="card-title">Gestión de pedidos (Operador / Admin)</h5>
            <p class="card-text">
              Aquí se ven todos los pedidos y se pueden cambiar estados:
            </p>
            <ul>
              <li>
                CREADO → ACEPTADO → EN_PREPARACIÓN → DESPACHADO → ENTREGADO /
                CANCELADO
              </li>
            </ul>
            <p class="mb-0">
              <strong>Regla:</strong>
              No se puede DESPACHAR sin ACEPTAR primero.
            </p>
          </div>
        </div>
      </div>
    </main>
  `,
  styles: [
    `
      :host {
        display: block;
        margin: 0;
        padding: 0;
        width: 100%;
        min-height: 100vh;
        background: #f8f9fa;
      }

      .top-navbar {
        background-color: #1e1e2d;
        width: 100%;
        height: 56px;
        display: flex;
        align-items: center;
        margin: 0;
        padding: 0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.25);
      }

      .nav-container {
        width: 100%;
        height: 100%;
        padding: 0 1.25rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }

      .nav-left {
        display: flex;
        align-items: center;
        gap: 0.25rem;
      }

      .nav-item {
        display: flex;
        align-items: center;
        gap: 0.4rem;
        color: #a2a3b7;
        text-decoration: none;
        font-size: 0.9rem;
        font-weight: 500;
        padding: 0.4rem 0.85rem;
        border-radius: 0.375rem;
        transition: all 0.2s ease;
        white-space: nowrap;
      }

      .nav-item:hover {
        color: #ffffff;
        background-color: rgba(255, 255, 255, 0.06);
      }

      .nav-item.active {
        color: #22c55e;
        font-weight: 600;
      }

      .nav-item i {
        font-size: 1rem;
      }

      .nav-right {
        display: flex;
        align-items: center;
        gap: 1rem;
      }

      .user-info {
        display: flex;
        align-items: center;
        gap: 0.6rem;
      }

      .avatar {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background-color: #3b82f6;
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        font-size: 0.85rem;
      }

      .user-text {
        line-height: 1.2;
      }

      .user-name {
        color: #ffffff;
        font-size: 0.85rem;
        font-weight: 500;
      }

      .user-role {
        color: #9ca3af;
        font-size: 0.7rem;
      }

      .btn-logout {
        background: transparent;
        border: 1px solid #4b5563;
        color: #e5e7eb;
        border-radius: 50px;
        padding: 0.3rem 0.9rem;
        font-size: 0.8rem;
        display: flex;
        align-items: center;
        gap: 0.35rem;
        cursor: pointer;
        transition: all 0.2s ease;
      }

      .btn-logout:hover {
        background-color: #ef4444;
        border-color: #ef4444;
        color: white;
      }

      .main-content {
        min-height: calc(100vh - 56px);
      }
    `,
  ],
})
export class OrdersComponent implements OnInit {
  nombre = "";
  roles: string[] = [];

  esAdmin = false;
  esOperador = false;
  esCliente = false;

  constructor(
    private auth: AuthService,
    private msal: MsalService,
    private cdr: ChangeDetectorRef,
  ) {}

  async ngOnInit() {
    // Datos para la barra (nombre + roles)
    const account = this.auth.getActiveAccount();
    this.nombre = account?.name || account?.username || "";

    const roles = await this.auth.getRoles();
    this.roles = roles;

    // Lógica original (no se tocó)
    this.esAdmin = roles.includes("ROLE_ADMINISTRADOR");
    this.esOperador = roles.includes("ROLE_OPERADOR");
    this.esCliente = roles.includes("ROLE_CLIENTE");

    console.log("Roles en Pedidos:", roles);
    console.log("Acceso Admin:", this.esAdmin);
    console.log("Acceso Operador:", this.esOperador);
    console.log("Acceso Cliente:", this.esCliente);

    this.cdr.detectChanges();
  }

  logout() {
    this.msal.logoutRedirect();
  }
}
