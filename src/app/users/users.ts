import { Component, OnInit, ChangeDetectorRef } from "@angular/core";
import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { FormsModule } from "@angular/forms";

import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";
import {
  UsuariosService,
  Usuario,
  CrearUsuarioPayload,
} from "../services/usuarios-service";

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
              Administra los usuarios de Microsoft Entra ID: crear, renombrar y
              cambiar roles.
            </p>
          </div>

          <button
            class="btn btn-primary"
            (click)="abrirModalCrear()"
            [disabled]="!esAdmin || cargando"
          >
            <i class="bi bi-person-plus me-2"></i>
            Nuevo usuario
          </button>
        </div>

        <!-- Solo admin puede ver el contenido -->
        <div *ngIf="!esAdmin" class="alert alert-warning">
          <i class="bi bi-shield-lock me-2"></i>
          Solo los administradores pueden gestionar usuarios.
        </div>

        <ng-container *ngIf="esAdmin">
          <!-- RESUMEN -->
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

          <!-- MENSAJES -->
          <div *ngIf="mensajeExito" class="alert alert-success alert-dismissible">
            {{ mensajeExito }}
            <button type="button" class="btn-close" (click)="mensajeExito = ''"></button>
          </div>
          <div *ngIf="mensajeError" class="alert alert-danger alert-dismissible">
            {{ mensajeError }}
            <button type="button" class="btn-close" (click)="mensajeError = ''"></button>
          </div>

          <!-- TABLA -->
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

                <div *ngIf="cargando" class="text-center py-4 text-muted">
                  <div class="spinner-border spinner-border-sm me-2"></div>
                  Cargando usuarios de Entra ID...
                </div>

                <div class="table-responsive" *ngIf="!cargando">
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
                            title="Editar nombre y rol"
                            (click)="abrirModalEditar(usuario)"
                          >
                            <i class="bi bi-pencil"></i>
                            Editar
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

            <div class="alert alert-info mt-4">
              <i class="bi bi-info-circle me-2"></i>
              <strong>Microsoft Entra ID:</strong>
              Los usuarios se gestionan directamente en Entra ID (crear, renombrar
              y rol). El rol se guarda en el campo <em>jobTitle</em>. Para que el
              token JWT incluya el rol de aplicación, asígnalo también como
              <strong>App role</strong> en el registro de la API en Azure Portal.
            </div>
          </div>
        </ng-container>
      </div>
    </main>

    <!-- ========== MODAL CREAR ========== -->
    <div class="modal-backdrop" *ngIf="mostrarModalCrear" (click)="cerrarModales()"></div>
    <div class="modal-panel" *ngIf="mostrarModalCrear">
      <div class="modal-header">
        <h5 class="mb-0">Nuevo usuario (Entra ID)</h5>
        <button type="button" class="btn-close" (click)="cerrarModales()"></button>
      </div>
      <div class="modal-body">
        <div class="mb-3">
          <label class="form-label">Nombre completo</label>
          <input
            type="text"
            class="form-control"
            [(ngModel)]="formCrear.nombre"
            placeholder="Ej: Juan Pérez"
            required
          />
        </div>
        <div class="mb-3">
          <label class="form-label">Correo (User Principal Name)</label>
          <input
            type="email"
            class="form-control"
            [(ngModel)]="formCrear.correo"
            placeholder="usuario@tudominio.com"
            required
          />
          <small class="text-muted">Debe ser un dominio verificado en tu tenant.</small>
        </div>
        <div class="mb-3">
          <label class="form-label">Rol</label>
          <select class="form-select" [(ngModel)]="formCrear.rol">
            <option value="ROLE_CLIENTE">Cliente</option>
            <option value="ROLE_OPERADOR">Operador</option>
            <option value="ROLE_ADMINISTRADOR">Administrador</option>
          </select>
        </div>
        <div class="mb-3">
          <label class="form-label">Contraseña temporal</label>
          <input
            type="text"
            class="form-control"
            [(ngModel)]="formCrear.passwordTemporal"
            placeholder="Mín. 8 caracteres, mayúscula, número y símbolo"
            required
          />
          <small class="text-muted">El usuario deberá cambiarla en el próximo inicio de sesión.</small>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" (click)="cerrarModales()">
          Cancelar
        </button>
        <button
          type="button"
          class="btn btn-primary"
          (click)="crearUsuario()"
          [disabled]="guardando || !formCrearValido()"
        >
          <span *ngIf="guardando" class="spinner-border spinner-border-sm me-1"></span>
          Crear usuario
        </button>
      </div>
    </div>

    <!-- ========== MODAL EDITAR ========== -->
    <div class="modal-backdrop" *ngIf="mostrarModalEditar" (click)="cerrarModales()"></div>
    <div class="modal-panel" *ngIf="mostrarModalEditar">
      <div class="modal-header">
        <h5 class="mb-0">Editar usuario</h5>
        <button type="button" class="btn-close" (click)="cerrarModales()"></button>
      </div>
      <div class="modal-body">
        <div class="mb-3">
          <label class="form-label">Correo</label>
          <input
            type="text"
            class="form-control"
            [value]="usuarioEditando?.correo"
            disabled
          />
        </div>
        <div class="mb-3">
          <label class="form-label">Nombre</label>
          <input
            type="text"
            class="form-control"
            [(ngModel)]="formEditar.nombre"
            required
          />
        </div>
        <div class="mb-3">
          <label class="form-label">Rol</label>
          <select class="form-select" [(ngModel)]="formEditar.rol">
            <option value="ROLE_CLIENTE">Cliente</option>
            <option value="ROLE_OPERADOR">Operador</option>
            <option value="ROLE_ADMINISTRADOR">Administrador</option>
          </select>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" (click)="cerrarModales()">
          Cancelar
        </button>
        <button
          type="button"
          class="btn btn-primary"
          (click)="guardarEdicion()"
          [disabled]="guardando || !formEditar.nombre.trim()"
        >
          <span *ngIf="guardando" class="spinner-border spinner-border-sm me-1"></span>
          Guardar cambios
        </button>
      </div>
    </div>
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

      .content-wrapper {
        max-width: 1100px;
        margin: 0 auto;
        padding: 1.5rem 1.25rem;
      }

      .summary-row {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
      }

      .summary-card {
        background: white;
        border-radius: 0.5rem;
        padding: 1.25rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
        display: flex;
        justify-content: space-between;
        align-items: center;
      }

      .table-section {
        width: 100%;
      }

      .modal-backdrop {
        position: fixed;
        inset: 0;
        background: rgba(0, 0, 0, 0.45);
        z-index: 1040;
      }

      .modal-panel {
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        border-radius: 0.5rem;
        width: 95%;
        max-width: 480px;
        z-index: 1050;
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
        display: flex;
        flex-direction: column;
      }

      .modal-header {
        padding: 1rem 1.25rem;
        border-bottom: 1px solid #e5e7eb;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }

      .modal-body {
        padding: 1.25rem;
      }

      .modal-footer {
        padding: 1rem 1.25rem;
        border-top: 1px solid #e5e7eb;
        display: flex;
        justify-content: flex-end;
        gap: 0.5rem;
      }

      @media (max-width: 768px) {
        .summary-row {
          grid-template-columns: 1fr;
        }
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

  usuarios: Usuario[] = [];
  usuariosFiltrados: Usuario[] = [];

  cargando = false;
  guardando = false;
  mensajeExito = "";
  mensajeError = "";

  mostrarModalCrear = false;
  formCrear = {
    nombre: "",
    correo: "",
    rol: "ROLE_CLIENTE",
    passwordTemporal: "",
  };

  mostrarModalEditar = false;
  usuarioEditando: Usuario | null = null;
  formEditar = {
    nombre: "",
    rol: "ROLE_CLIENTE",
  };

  constructor(
    private auth: AuthService,
    private msal: MsalService,
    private cdr: ChangeDetectorRef,
    private usuariosService: UsuariosService,
  ) {}

  async ngOnInit() {
    const account = this.auth.getActiveAccount();
    this.nombre = account?.name || account?.username || "Administrador";

    const roles = await this.auth.getRoles();
    this.roles = roles;
    this.esAdmin = roles.includes("ROLE_ADMINISTRADOR");
    this.esOperador = roles.includes("ROLE_OPERADOR");

    if (this.esAdmin) {
      this.cargarUsuarios();
    }

    this.cdr.detectChanges();
  }

  cargarUsuarios() {
    this.cargando = true;
    this.mensajeError = "";

    this.usuariosService.listarUsuarios().subscribe({
      next: (lista: Usuario[]) => {
        this.usuarios = lista;
        this.filtrarUsuarios();
        this.cargando = false;
        this.cdr.detectChanges();
      },
      error: (err: any) => {
        console.error(err);
        this.mensajeError =
          "No se pudieron cargar los usuarios de Entra ID. Verifica permisos Graph (User.ReadWrite.All) y consentimiento de admin.";
        this.cargando = false;
        this.cdr.detectChanges();
      },
    });
  }

  get usuariosActivos(): number {
    return this.usuarios.filter((u) => u.estado === "ACTIVO").length;
  }

  get administradores(): number {
    return this.usuarios.filter((u) => u.rol === "ROLE_ADMINISTRADOR").length;
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

  abrirModalCrear() {
    this.formCrear = {
      nombre: "",
      correo: "",
      rol: "ROLE_CLIENTE",
      passwordTemporal: "",
    };
    this.mostrarModalCrear = true;
    this.mensajeError = "";
  }

  formCrearValido(): boolean {
    return (
      this.formCrear.nombre.trim().length > 0 &&
      this.formCrear.correo.trim().includes("@") &&
      this.formCrear.passwordTemporal.length >= 8
    );
  }

  crearUsuario() {
    if (!this.formCrearValido()) return;

    this.guardando = true;
    this.mensajeError = "";
    this.mensajeExito = "";

    const payload: CrearUsuarioPayload = {
      nombre: this.formCrear.nombre.trim(),
      correo: this.formCrear.correo.trim(),
      rol: this.formCrear.rol,
      passwordTemporal: this.formCrear.passwordTemporal,
    };

    this.usuariosService.crearUsuario(payload).subscribe({
      next: (usuario: Usuario | null) => {
        this.guardando = false;
        if (usuario) {
          this.usuarios = [usuario, ...this.usuarios];
          this.filtrarUsuarios();
          this.mensajeExito = `Usuario "${usuario.nombre}" creado correctamente en Entra ID.`;
          this.cerrarModales();
        } else {
          this.mensajeError =
            "No se pudo crear el usuario. Revisa permisos y el dominio del correo.";
        }
        this.cdr.detectChanges();
      },
      error: (err: any) => {
        this.guardando = false;
        const detail =
          err?.error?.error?.message ||
          err?.message ||
          "Error desconocido al crear usuario.";
        this.mensajeError = `Error al crear: ${detail}`;
        this.cdr.detectChanges();
      },
    });
  }

  abrirModalEditar(usuario: Usuario) {
    this.usuarioEditando = usuario;
    this.formEditar = {
      nombre: usuario.nombre,
      rol: usuario.rol,
    };
    this.mostrarModalEditar = true;
    this.mensajeError = "";
  }

  guardarEdicion() {
    if (!this.usuarioEditando || !this.formEditar.nombre.trim()) return;

    this.guardando = true;
    this.mensajeError = "";
    this.mensajeExito = "";

    this.usuariosService
      .actualizarUsuario(this.usuarioEditando.id, {
        nombre: this.formEditar.nombre.trim(),
        rol: this.formEditar.rol,
      })
      .subscribe({
        next: (actualizado: Usuario | null) => {
          this.guardando = false;
          if (actualizado) {
            const idx = this.usuarios.findIndex((u) => u.id === actualizado.id);
            if (idx !== -1) {
              this.usuarios[idx] = actualizado;
            }
            this.filtrarUsuarios();
            this.mensajeExito = `Usuario "${actualizado.nombre}" actualizado correctamente.`;
            this.cerrarModales();
          } else {
            this.mensajeError = "No se pudo actualizar el usuario.";
          }
          this.cdr.detectChanges();
        },
        error: (err: any) => {
          this.guardando = false;
          const detail =
            err?.error?.error?.message ||
            err?.message ||
            "Error desconocido al actualizar.";
          this.mensajeError = `Error al actualizar: ${detail}`;
          this.cdr.detectChanges();
        },
      });
  }

  cerrarModales() {
    this.mostrarModalCrear = false;
    this.mostrarModalEditar = false;
    this.usuarioEditando = null;
  }

  logout() {
    this.msal.logoutRedirect();
  }
}
