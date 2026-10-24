import { Component, OnInit } from "@angular/core";
import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";
import { ChangeDetectorRef } from "@angular/core";
import { PedidoService } from "../services/pedido-service";
import { Pedido } from "../services/pedido";

@Component({
  selector: "app-dashboard",
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <!-- ========== NAVBAR SUPERIOR ========== -->
    <nav class="top-navbar" aria-label="Navegación principal">
      <div class="nav-container">
        <!-- Menú horizontal izquierda -->
        <div class="nav-left">
          <a class="nav-item active" routerLink="/dashboard">
            <i class="bi bi-house-door-fill"></i>
            <span>Dashboard</span>
          </a>

          <a class="nav-item" routerLink="/orders">
            <i class="bi bi-cart3"></i>
            <span>Pedidos</span>
          </a>

          <a
            class="nav-item"
            *ngIf="esAdmin || esOperador"
            routerLink="/catalog"
          >
            <i class="bi bi-journal-text"></i>
            <span>Catálogo</span>
          </a>

          <!-- Usuarios (solo Admin) -->
          <a class="nav-item" *ngIf="esAdmin" routerLink="/usuarios">
            <i class="bi bi-people"></i>
            <span>Usuarios</span>
          </a>

          <a class="nav-item" href="javascript:void(0)">
            <i class="bi bi-box"></i>
            <span>TEST</span>
          </a>
        </div>

        <!-- Usuario + Logout derecha -->
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
            <span>Cerrar sesión</span>
          </button>
        </div>
      </div>
    </nav>

    <!-- ========== CONTENIDO ========== -->
    <main class="main-content">
      <div class="content-wrapper">
        <!-- BIENVENIDA -->
        <header class="welcome">
          <h1>Hola, {{ nombre.split(" ")[0] || "Usuario" }}</h1>
          <p *ngIf="roles.length > 0">
            Estas son las secciones de la aplicación y lo que puedes hacer en
            cada una con tu cuenta.
          </p>
        </header>

        <!-- Sin roles -->
        <div *ngIf="roles.length === 0" class="notice" role="status">
          <i class="bi bi-info-circle-fill"></i>
          <p>
            Tu cuenta todavía no tiene roles asignados. Pide a un administrador
            que te asigne uno para acceder a las secciones.
          </p>
        </div>

        <!-- ACCESOS -->
        <section class="tiles" aria-label="Secciones de la aplicación">
          <a class="tile" routerLink="/orders">
            <span class="tile-icon tone-green"><i class="bi bi-cart3"></i></span>
            <span class="tile-text">
              <strong>Pedidos</strong>
              <small>Consulta y gestiona los pedidos.</small>
            </span>
            <i class="bi bi-chevron-right tile-go"></i>
          </a>

          <a
            class="tile"
            *ngIf="esAdmin || esOperador"
            routerLink="/catalog"
          >
            <span class="tile-icon tone-amber"><i class="bi bi-journal-text"></i></span>
            <span class="tile-text">
              <strong>Catálogo</strong>
              <small>Revisa y edita los productos.</small>
            </span>
            <i class="bi bi-chevron-right tile-go"></i>
          </a>
          <div
            class="tile is-locked"
            *ngIf="!(esAdmin || esOperador)"
            aria-disabled="true"
          >
            <span class="tile-icon"><i class="bi bi-lock"></i></span>
            <span class="tile-text">
              <strong>Catálogo</strong>
              <small>Disponible para operadores y administradores.</small>
            </span>
          </div>

          <a class="tile" *ngIf="esAdmin" routerLink="/usuarios">
            <span class="tile-icon tone-rose"><i class="bi bi-people"></i></span>
            <span class="tile-text">
              <strong>Usuarios</strong>
              <small>Crea cuentas y asigna roles.</small>
            </span>
            <i class="bi bi-chevron-right tile-go"></i>
          </a>
          <div class="tile is-locked" *ngIf="!esAdmin" aria-disabled="true">
            <span class="tile-icon"><i class="bi bi-lock"></i></span>
            <span class="tile-text">
              <strong>Usuarios</strong>
              <small>Solo para administradores.</small>
            </span>
          </div>
        </section>

        <!-- SESIÓN Y VISTAS -->
        <div class="split">
          <section class="panel session" aria-labelledby="titulo-sesion">
            <h2 id="titulo-sesion" class="panel-title">Tu sesión</h2>

            <div class="session-user">
              <div class="session-avatar">
                {{ nombre ? nombre.charAt(0).toUpperCase() : "U" }}
              </div>
              <div class="session-user-text">
                <span class="session-label">Usuario</span>
                <strong>{{ nombre || "Usuario" }}</strong>
              </div>
            </div>

            <div class="session-roles">
              <span class="session-label">Roles</span>
              <div class="chips">
                <ng-container *ngFor="let rol of roles">
                  <span class="chip" [attr.data-rol]="rol" [ngSwitch]="rol">
                    <ng-container *ngSwitchCase="'ROLE_ADMINISTRADOR'">Administrador</ng-container>
                    <ng-container *ngSwitchCase="'ROLE_OPERADOR'">Operador</ng-container>
                    <ng-container *ngSwitchCase="'ROLE_CLIENTE'">Cliente</ng-container>
                    <ng-container *ngSwitchDefault>{{ rol }}</ng-container>
                  </span>
                </ng-container>
                <span *ngIf="roles.length === 0" class="chip chip-none">Sin roles</span>
              </div>
            </div>
          </section>

          <section class="panel" aria-labelledby="titulo-vistas">
            <h2 id="titulo-vistas" class="panel-title">Qué puedes hacer</h2>

            <div class="role-views">
              <div *ngIf="esAdmin" class="role-view" data-rol="ROLE_ADMINISTRADOR">
                <span class="role-view-icon"><i class="bi bi-shield-lock"></i></span>
                <div>
                  <h3>Vista Administrador</h3>
                  <p>Puedes ver todo el sistema, pedidos y catálogo.</p>
                </div>
              </div>

              <div *ngIf="esOperador" class="role-view" data-rol="ROLE_OPERADOR">
                <span class="role-view-icon"><i class="bi bi-tools"></i></span>
                <div>
                  <h3>Vista Operador</h3>
                  <p>
                    Gestionas pedidos y estados. También puedes ver/editar
                    catálogo.
                  </p>
                </div>
              </div>

              <div *ngIf="esCliente" class="role-view" data-rol="ROLE_CLIENTE">
                <span class="role-view-icon"><i class="bi bi-person"></i></span>
                <div>
                  <h3>Vista Cliente</h3>
                  <p>
                    Puedes crear y ver tus propios pedidos. Solo consulta
                    catálogo.
                  </p>
                </div>
              </div>

              <p *ngIf="!esAdmin && !esOperador && !esCliente" class="role-empty">
                Cuando te asignen un rol, aquí verás qué puedes hacer con él.
              </p>
            </div>
          </section>
        </div>
      </div>
    </main>
  `,
  styles: [
    `
      :host {
        --ink: #1e1e2d;
        --canvas: #f3f4f8;
        --surface: #ffffff;
        --line: #e3e5ed;
        --text: #1f2233;
        --muted: #667085;
        --brand: #15803d;
        --brand-soft: #e7f6ec;
        --focus: #2563eb;

        display: block;
        margin: 0;
        padding: 0;
        width: 100%;
        min-height: 100vh;
        background: var(--canvas);
        color: var(--text);
        font-family: "Segoe UI Variable", "Segoe UI", system-ui, -apple-system,
          Roboto, "Helvetica Neue", Arial, sans-serif;
      }

      :host *:focus-visible {
        outline: 2px solid var(--focus);
        outline-offset: 2px;
      }

      /* ---------- Navbar ---------- */
      .top-navbar {
        position: sticky;
        top: 0;
        z-index: 1000;
        background-color: var(--ink);
        width: 100%;
        height: 58px;
        display: flex;
        align-items: center;
        box-shadow: 0 1px 0 rgba(255, 255, 255, 0.05), 0 4px 14px rgba(15, 16, 30, 0.25);
      }

      .nav-container {
        width: 100%;
        height: 100%;
        padding: 0 1.25rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
      }

      .nav-left {
        display: flex;
        align-items: stretch;
        height: 100%;
        gap: 0.25rem;
        overflow-x: auto;
        scrollbar-width: none;
      }

      .nav-left::-webkit-scrollbar {
        display: none;
      }

      .nav-item {
        position: relative;
        display: flex;
        align-items: center;
        gap: 0.45rem;
        color: #a9aabf;
        text-decoration: none;
        font-size: 0.9rem;
        font-weight: 500;
        padding: 0 0.9rem;
        transition: color 0.15s ease, background-color 0.15s ease;
        white-space: nowrap;
      }

      .nav-item:hover {
        color: #ffffff;
        background-color: rgba(255, 255, 255, 0.05);
      }

      .nav-item.active {
        color: #4ade80;
        font-weight: 600;
      }

      .nav-item.active::after {
        content: "";
        position: absolute;
        left: 0.9rem;
        right: 0.9rem;
        bottom: 0;
        height: 3px;
        border-radius: 3px 3px 0 0;
        background: #22c55e;
      }

      .nav-item i {
        font-size: 1.05rem;
      }

      .nav-right {
        display: flex;
        align-items: center;
        gap: 1rem;
        flex-shrink: 0;
      }

      .user-info {
        display: flex;
        align-items: center;
        gap: 0.65rem;
      }

      .avatar {
        width: 34px;
        height: 34px;
        border-radius: 50%;
        background-color: #3b82f6;
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        font-size: 0.88rem;
      }

      .user-text {
        line-height: 1.2;
      }

      .user-name {
        color: #ffffff;
        font-size: 0.85rem;
        font-weight: 600;
      }

      .user-role {
        color: #a3a5ba;
        font-size: 0.72rem;
      }

      .btn-logout {
        background: transparent;
        border: 1px solid #484a63;
        color: #e5e7eb;
        border-radius: 999px;
        padding: 0.35rem 0.95rem;
        font-size: 0.8rem;
        display: flex;
        align-items: center;
        gap: 0.4rem;
        cursor: pointer;
        transition: all 0.15s ease;
      }

      .btn-logout:hover {
        background-color: #ef4444;
        border-color: #ef4444;
        color: white;
      }

      /* ---------- Layout ---------- */
      .main-content {
        min-height: calc(100vh - 58px);
      }

      .content-wrapper {
        max-width: 1100px;
        margin: 0 auto;
        padding: 2.25rem 1.25rem 3rem;
      }

      .welcome {
        margin-bottom: 1.75rem;
      }

      .welcome h1 {
        margin: 0 0 0.35rem;
        font-size: 2rem;
        font-weight: 700;
        letter-spacing: -0.025em;
      }

      .welcome p {
        margin: 0;
        max-width: 56ch;
        color: var(--muted);
        font-size: 0.97rem;
        line-height: 1.55;
      }

      .notice {
        display: flex;
        gap: 0.7rem;
        padding: 0.9rem 1rem;
        margin-bottom: 1.5rem;
        background: #eef4ff;
        border: 1px solid #d3e2ff;
        border-radius: 12px;
        color: #243b6b;
        font-size: 0.92rem;
        line-height: 1.5;
      }

      .notice i {
        margin-top: 0.15rem;
        color: #2f55b8;
      }

      .notice p {
        margin: 0;
      }

      /* ---------- Accesos ---------- */
      .tiles {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
        margin-bottom: 1.5rem;
      }

      .tile {
        display: flex;
        align-items: center;
        gap: 0.95rem;
        padding: 1.15rem 1.2rem;
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 14px;
        color: inherit;
        text-decoration: none;
        transition: border-color 0.15s ease, box-shadow 0.15s ease;
      }

      a.tile:hover {
        border-color: #9fd3b2;
        box-shadow: 0 4px 16px rgba(21, 128, 61, 0.1);
      }

      .tile-icon {
        width: 46px;
        height: 46px;
        flex-shrink: 0;
        border-radius: 13px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        background: #eceef4;
        color: #8a90a5;
      }

      .tone-green {
        background: var(--brand-soft);
        color: var(--brand);
      }

      .tone-amber {
        background: #fdf1d8;
        color: #8a5a00;
      }

      .tone-rose {
        background: #fde8ec;
        color: #b4233f;
      }

      .tile-text {
        display: flex;
        flex-direction: column;
        gap: 0.15rem;
        min-width: 0;
        flex: 1;
      }

      .tile-text strong {
        font-size: 1rem;
      }

      .tile-text small {
        font-size: 0.83rem;
        line-height: 1.4;
        color: var(--muted);
      }

      .tile-go {
        color: #aab0c4;
        transition: transform 0.15s ease, color 0.15s ease;
      }

      a.tile:hover .tile-go {
        color: var(--brand);
        transform: translateX(3px);
      }

      .tile.is-locked {
        background: transparent;
        border-style: dashed;
        border-color: #cfd2de;
        cursor: not-allowed;
      }

      .tile.is-locked strong {
        color: #6b7087;
      }

      /* ---------- Paneles ---------- */
      .split {
        display: grid;
        grid-template-columns: 1fr 2fr;
        gap: 1rem;
        align-items: start;
      }

      .panel {
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 1.3rem 1.35rem;
      }

      .panel-title {
        margin: 0 0 1.1rem;
        font-size: 1.05rem;
        font-weight: 700;
      }

      .session-user {
        display: flex;
        align-items: center;
        gap: 0.85rem;
        padding-bottom: 1.1rem;
        margin-bottom: 1.1rem;
        border-bottom: 1px solid var(--line);
      }

      .session-avatar {
        width: 48px;
        height: 48px;
        flex-shrink: 0;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: #3b82f6;
        color: #fff;
        font-weight: 700;
        font-size: 1.15rem;
      }

      .session-user-text {
        display: flex;
        flex-direction: column;
        min-width: 0;
      }

      .session-user-text strong {
        overflow-wrap: anywhere;
      }

      .session-label {
        display: block;
        font-size: 0.8rem;
        color: var(--muted);
        margin-bottom: 0.1rem;
      }

      .session-roles .session-label {
        margin-bottom: 0.5rem;
      }

      .chips {
        display: flex;
        flex-wrap: wrap;
        gap: 0.4rem;
      }

      .chip {
        display: inline-flex;
        align-items: center;
        padding: 0.25rem 0.75rem;
        border-radius: 999px;
        font-size: 0.8rem;
        font-weight: 600;
        background: #eceef4;
        color: #4a4f66;
      }

      .chip[data-rol="ROLE_ADMINISTRADOR"] {
        background: #fde8ec;
        color: #b4233f;
      }

      .chip[data-rol="ROLE_OPERADOR"] {
        background: #fdf1d8;
        color: #8a5a00;
      }

      .chip[data-rol="ROLE_CLIENTE"] {
        background: #e4edff;
        color: #2f55b8;
      }

      .role-views {
        display: flex;
        flex-direction: column;
        gap: 0.75rem;
      }

      .role-view {
        display: flex;
        gap: 0.95rem;
        padding: 1rem 1.1rem;
        border-radius: 12px;
        border-left: 4px solid;
      }

      .role-view h3 {
        margin: 0 0 0.2rem;
        font-size: 0.98rem;
        font-weight: 700;
      }

      .role-view p {
        margin: 0;
        font-size: 0.9rem;
        line-height: 1.5;
        color: var(--text);
      }

      .role-view-icon {
        width: 38px;
        height: 38px;
        flex-shrink: 0;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.1rem;
        background: #fff;
      }

      .role-view[data-rol="ROLE_ADMINISTRADOR"] {
        background: #fef5f7;
        border-color: #d9435f;
      }

      .role-view[data-rol="ROLE_ADMINISTRADOR"] .role-view-icon {
        color: #b4233f;
      }

      .role-view[data-rol="ROLE_OPERADOR"] {
        background: #fffaf0;
        border-color: #e9a319;
      }

      .role-view[data-rol="ROLE_OPERADOR"] .role-view-icon {
        color: #8a5a00;
      }

      .role-view[data-rol="ROLE_CLIENTE"] {
        background: #f5f8ff;
        border-color: #4f7bde;
      }

      .role-view[data-rol="ROLE_CLIENTE"] .role-view-icon {
        color: #2f55b8;
      }

      .role-empty {
        margin: 0;
        color: var(--muted);
        font-size: 0.92rem;
      }

      @media (prefers-reduced-motion: reduce) {
        .tile-go {
          transition: none;
        }

        a.tile:hover .tile-go {
          transform: none;
        }
      }

      /* ---------- Responsive ---------- */
      @media (max-width: 900px) {
        .user-text {
          display: none;
        }

        .tiles {
          grid-template-columns: 1fr;
        }

        .split {
          grid-template-columns: 1fr;
        }
      }

      @media (max-width: 600px) {
        .welcome h1 {
          font-size: 1.6rem;
        }

        .btn-logout span {
          display: none;
        }

        .btn-logout {
          padding: 0.45rem 0.65rem;
        }

        .nav-item {
          padding: 0 0.65rem;
        }
      }
    `,
  ],
})
export class DashboardComponent implements OnInit {
  nombre = "";
  roles: string[] = [];
  esAdmin = false;
  esOperador = false;
  esCliente = false;
  pedidos: Pedido[] = [];

  constructor(
    private auth: AuthService,
    private msal: MsalService,
    private cdr: ChangeDetectorRef,
    private pedidoService: PedidoService,
  ) {}

  async ngOnInit() {
    const account = this.auth.getActiveAccount();
    this.nombre = account?.name || account?.username || "";

    const roles = await this.auth.getRoles();
    this.roles = roles;

    this.esAdmin = roles.includes("ROLE_ADMINISTRADOR");
    this.esOperador = roles.includes("ROLE_OPERADOR");
    this.esCliente = roles.includes("ROLE_CLIENTE");

    console.log("Roles del Dashboard:", this.roles);

    this.pedidoService.listarPedidos().subscribe({
      next: (pedidos) => {
        this.pedidos = pedidos;
        console.log("Pedidos cargados:", this.pedidos);
      },
      error: (error) => {
        console.error("Error obteniendo pedidos:", error);
      },
    });

    this.cdr.detectChanges();
  }

  logout() {
    this.msal.logoutRedirect();
  }
}