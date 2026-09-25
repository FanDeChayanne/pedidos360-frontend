import { Component, OnInit, ChangeDetectorRef } from "@angular/core";

import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { FormsModule } from "@angular/forms";

import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";

interface UsuarioMock {
  nombre: string;
  correo: string;
  rol: string;
  estado: string;
}

@Component({
  selector: "app-usuarios",
  standalone: true,
  imports: [CommonModule, RouterLink, FormsModule],
  template: `
    <!-- ========== NAVBAR ========== -->
    <nav class="top-navbar">
      <div class="nav-container">
        <div class="nav-left">
          <a class="nav-item" routerLink="/dashboard">
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

          <a class="nav-item active" *ngIf="esAdmin" routerLink="/usuarios">
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

    <!-- ========== CONTENIDO ========== -->
    <main class="main-content">
      <div class="content-wrapper">
        <!-- ENCABEZADO -->
        <div class="d-flex justify-content-between align-items-center mb-4">
          <div>
            <h1 class="fw-bold mb-1">Gestión de usuarios</h1>
            <p class="text-muted mb-0">
              Administra los usuarios y sus roles dentro del sistema.
            </p>
          </div>

          <button class="btn btn-primary">
            <i class="bi bi-person-plus me-2"></i>
            Nuevo usuario
          </button>
        </div>

        <!-- RESUMEN EN LÍNEA -->
        <div class="summary-row mb-4">
          <div class="summary-card">
            <div>
              <p class="text-muted mb-1">Total usuarios</p>
              <h3 class="fw-bold mb-0">{{ usuarios.length }}</h3>
            </div>
            <div class="text-primary fs-2">
              <i class="bi bi-people"></i>
            </div>
          </div>

          <div class="summary-card">
            <div>
              <p class="text-muted mb-1">Usuarios activos</p>
              <h3 class="fw-bold mb-0">{{ usuariosActivos }}</h3>
            </div>
            <div class="text-success fs-2">
              <i class="bi bi-person-check"></i>
            </div>
          </div>

          <div class="summary-card">
            <div>
              <p class="text-muted mb-1">Administradores</p>
              <h3 class="fw-bold mb-0">{{ administradores }}</h3>
            </div>
            <div class="text-danger fs-2">
              <i class="bi bi-shield-lock"></i>
            </div>
          </div>
        </div>

        <!-- ========== SECCIÓN DE USUARIOS CENTRADA ========== -->
        <div class="table-section">
          <div class="card border-0 shadow-sm">
            <div class="card-body">
              <!-- FILTROS -->
              <div class="row g-3 mb-4">
                <div class="col-md-8">
                  <div class="input-group">
                    <span class="input-group-text">
                      <i class="bi bi-search"></i>
                    </span>
                    <input
                      type="text"
                      class="form-control"
                      placeholder="Buscar por nombre o correo..."
                      [(ngModel)]="busqueda"
                      (input)="filtrarUsuarios()"
                    />
                  </div>
                </div>

                <div class="col-md-4">
                  <select
                    class="form-select"
                    [(ngModel)]="filtroRol"
                    (change)="filtrarUsuarios()"
                  >
                    <option value="TODOS">Todos los roles</option>
                    <option value="ROLE_ADMINISTRADOR">Administrador</option>
                    <option value="ROLE_OPERADOR">Operador</option>
                    <option value="ROLE_CLIENTE">Cliente</option>
                  </select>
                </div>
              </div>

              <!-- TABLA -->
              <div class="table-responsive">
                <table class="table table-hover align-middle mb-0">
                  <thead class="table-light">
                    <tr>
                      <th>Usuario</th>
                      <th>Correo</th>
                      <th>Rol</th>
                      <th>Estado</th>
                      <th class="text-end">Acciones</th>
                    </tr>
                  </thead>

                  <tbody>
                    <tr *ngFor="let usuario of usuariosFiltrados">
                      <td>
                        <div class="d-flex align-items-center">
                          <div
                            class="bg-primary text-white rounded-circle d-flex align-items-center justify-content-center me-2"
                            style="width:40px; height:40px; flex-shrink: 0;"
                          >
                            <i class="bi bi-person"></i>
                          </div>
                          <span class="fw-semibold">{{ usuario.nombre }}</span>
                        </div>
                      </td>

                      <td class="text-muted">{{ usuario.correo }}</td>

                      <td>
                        <span
                          *ngIf="usuario.rol === 'ROLE_ADMINISTRADOR'"
                          class="badge text-bg-danger"
                        >
                          Administrador
                        </span>
                        <span
                          *ngIf="usuario.rol === 'ROLE_OPERADOR'"
                          class="badge text-bg-warning"
                        >
                          Operador
                        </span>
                        <span
                          *ngIf="usuario.rol === 'ROLE_CLIENTE'"
                          class="badge text-bg-primary"
                        >
                          Cliente
                        </span>
                      </td>

                      <td>
                        <span
                          *ngIf="usuario.estado === 'ACTIVO'"
                          class="badge text-bg-success"
                        >
                          Activo
                        </span>
                        <span
                          *ngIf="usuario.estado === 'INACTIVO'"
                          class="badge text-bg-secondary"
                        >
                          Inactivo
                        </span>
                      </td>

                      <td class="text-end">
                        <button
                          class="btn btn-sm btn-outline-primary me-2"
                          title="Editar usuario"
                        >
                          <i class="bi bi-pencil"></i>
                        </button>
                        <button
                          class="btn btn-sm btn-outline-secondary"
                          title="Cambiar estado"
                        >
                          <i class="bi bi-person-x"></i>
                        </button>
                      </td>
                    </tr>

                    <tr *ngIf="usuariosFiltrados.length === 0">
                      <td colspan="5" class="text-center py-5 text-muted">
                        <i class="bi bi-search fs-2 d-block mb-2"></i>
                        No se encontraron usuarios.
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <!-- AVISO -->
          <div class="alert alert-info mt-4">
            <i class="bi bi-info-circle me-2"></i>
            <strong>Modo de prueba:</strong>
            los usuarios mostrados actualmente son datos simulados. La gestión
            real se conectará al backend cuando esté disponible.
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
        padding: 2rem 1rem;
      }

      .content-wrapper {
        max-width: 1100px;
        margin: 0 auto;
      }

      /* Tarjetas de resumen */
      .summary-row {
        display: flex;
        gap: 1rem;
      }

      .summary-card {
        flex: 1;
        background: white;
        border-radius: 0.5rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
        padding: 1.25rem 1.5rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }

      /* ===== ESTA ES LA SECCIÓN QUE SE CENTRA ===== */
      .table-section {
        width: 100%;
        margin: 0 auto;
      }

      .table-section .table {
        width: 100%;
      }
    `,
  ],
})
export class UsuariosComponent implements OnInit {
  nombre = "";
  roles: string[] = [];
  esAdmin = false;
  esOperador = false;

  busqueda = "";
  filtroRol = "TODOS";

  usuarios: UsuarioMock[] = [
    {
      nombre: "Administrador de prueba",
      correo: "admin@pedidos360.cl",
      rol: "ROLE_ADMINISTRADOR",
      estado: "ACTIVO",
    },
    {
      nombre: "Operador de prueba",
      correo: "operador@pedidos360.cl",
      rol: "ROLE_OPERADOR",
      estado: "ACTIVO",
    },
    {
      nombre: "Cliente de prueba",
      correo: "cliente@pedidos360.cl",
      rol: "ROLE_CLIENTE",
      estado: "ACTIVO",
    },
    {
      nombre: "Cliente inactivo",
      correo: "cliente2@pedidos360.cl",
      rol: "ROLE_CLIENTE",
      estado: "INACTIVO",
    },
  ];

  usuariosFiltrados: UsuarioMock[] = [];

  constructor(
    private auth: AuthService,
    private msal: MsalService,
    private cdr: ChangeDetectorRef,
  ) {}

  async ngOnInit() {
    const account = this.auth.getActiveAccount();
    this.nombre = account?.name || account?.username || "Administrador";

    const roles = await this.auth.getRoles();
    this.roles = roles;
    this.esAdmin = roles.includes("ROLE_ADMINISTRADOR");
    this.esOperador = roles.includes("ROLE_OPERADOR");

    this.usuariosFiltrados = [...this.usuarios];

    this.cdr.detectChanges();
  }

  logout() {
    this.msal.logoutRedirect();
  }

  get usuariosActivos(): number {
    return this.usuarios.filter((usuario) => usuario.estado === "ACTIVO")
      .length;
  }

  get administradores(): number {
    return this.usuarios.filter(
      (usuario) => usuario.rol === "ROLE_ADMINISTRADOR",
    ).length;
  }

  filtrarUsuarios(): void {
    const texto = this.busqueda.toLowerCase().trim();

    this.usuariosFiltrados = this.usuarios.filter((usuario) => {
      const coincideTexto =
        usuario.nombre.toLowerCase().includes(texto) ||
        usuario.correo.toLowerCase().includes(texto);

      const coincideRol =
        this.filtroRol === "TODOS" || usuario.rol === this.filtroRol;

      return coincideTexto && coincideRol;
    });
  }
}
