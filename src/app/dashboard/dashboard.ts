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
    <!-- ========== NAVBAR SUPERIOR (sin espacios) ========== -->
    <nav class="top-navbar">
      <div class="nav-container">
        <!-- Menú horizontal izquierda -->
        <div class="nav-left">
          <a class="nav-item active" routerLink="/dashboard">
            <i class="bi bi-house-door-fill"></i>
            Dashboard
          </a>

          <a class="nav-item" routerLink="/orders">
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
            Cerrar sesión
          </button>
        </div>
      </div>
    </nav>

    <!-- ========== CONTENIDO ========== -->
    <main class="main-content">
      <div class="p-4">
        <div class="card border-0 shadow-sm mb-4">
          <div class="card-body">
            <h5 class="card-title mb-3">Información de sesión</h5>
            <p class="mb-1"><strong>Usuario:</strong> {{ nombre }}</p>
            <p class="mb-0">
              <strong>Roles:</strong> {{ roles.join(", ") || "Sin roles" }}
            </p>
          </div>
        </div>

        <div
          *ngIf="esAdmin"
          class="card border-0 shadow-sm mb-3"
          style="border-left: 4px solid #198754 !important;"
        >
          <div class="card-body">
            <h5 class="card-title text-success">Vista Administrador</h5>
            <p class="card-text mb-0">
              Puedes ver todo el sistema, pedidos y catálogo.
            </p>
          </div>
        </div>

        <div
          *ngIf="esOperador"
          class="card border-0 shadow-sm mb-3"
          style="border-left: 4px solid #fd7e14 !important;"
        >
          <div class="card-body">
            <h5 class="card-title text-warning">Vista Operador</h5>
            <p class="card-text mb-0">
              Gestionas pedidos y estados. También puedes ver/editar catálogo.
            </p>
          </div>
        </div>

        <div
          *ngIf="esCliente"
          class="card border-0 shadow-sm mb-3"
          style="border-left: 4px solid #0d6efd !important;"
        >
          <div class="card-body">
            <h5 class="card-title text-primary">Vista Cliente</h5>
            <p class="card-text mb-0">
              Puedes crear y ver tus propios pedidos. Solo consulta catálogo.
            </p>
          </div>
        </div>

        <!-- ========== PEDIDOS ========== -->
        <div class="card border-0 shadow-sm mb-4">
          <div class="card-body">
            <h5 class="card-title mb-3">Pedidos</h5>
            <div *ngIf="pedidos.length === 0" class="text-muted">
              No hay pedidos registrados.
            </div>
            <div *ngIf="pedidos.length > 0" class="table-responsive">
              <table class="table table-hover align-middle mb-0">
                <thead>
                  <tr>
                    <th>ID</th>
                    <th>Cliente</th>
                    <th>Producto</th>
                    <th>Cantidad</th>
                    <th>Estado</th>
                  </tr>
                </thead>
                <tbody>
                  <tr *ngFor="let pedido of pedidos">
                    <td>{{ pedido.id }}</td>
                    <td>{{ pedido.clienteId }}</td>
                    <td>{{ pedido.productoId }}</td>
                    <td>{{ pedido.cantidad }}</td>
                    <td>{{ pedido.estado }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </main>
  `,
  styles: [
    `
      /* ===== ELIMINAR TODOS LOS ESPACIOS BLANCOS ===== */
      :host {
        display: block;
        margin: 0 !important;
        padding: 0 !important;
        width: 100%;
        min-height: 100vh;
        background-color: #f8f9fa;
      }

      /* Forzar que no haya márgenes en el body desde este componente */
      :host ::ng-deep body,
      :host ::ng-deep html {
        margin: 0 !important;
        padding: 0 !important;
      }

      /* ===== NAVBAR ===== */
      .top-navbar {
        background-color: #1e1e2d;
        width: 100%;
        height: 56px;
        display: flex;
        align-items: center;
        margin: 0;
        padding: 0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.25);
        position: relative;
        z-index: 1000;
      }

      .nav-container {
        width: 100%;
        height: 100%;
        padding: 0 1.25rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }

      /* Menú izquierda - HORIZONTAL */
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

      /* Derecha */
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
        margin: 0;
        padding: 0;
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
