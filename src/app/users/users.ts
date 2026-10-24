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
    <nav class="top-navbar" aria-label="Navegación principal">
      <div class="nav-container">
        <div class="nav-left">
          <a class="nav-item" routerLink="/dashboard">
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

          <a class="nav-item active" *ngIf="esAdmin" routerLink="/usuarios">
            <i class="bi bi-people"></i>
            <span>Usuarios</span>
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
            <span>Cerrar sesión</span>
          </button>
        </div>
      </div>
    </nav>

    <!-- ========== CONTENIDO ========== -->
    <main class="main-content">
      <div class="content-wrapper">
        <!-- ENCABEZADO -->
        <header class="page-head">
          <div class="page-head-text">
            <h1>Gestión de usuarios</h1>
            <p>
              Administra los usuarios de Microsoft Entra ID: crea cuentas,
              cambia nombres y asigna roles.
            </p>
          </div>

          <button
            class="btn-brand"
            (click)="abrirModalCrear()"
            [disabled]="!esAdmin || cargando"
          >
            <i class="bi bi-person-plus"></i>
            Nuevo usuario
          </button>
        </header>

        <!-- Solo admin puede ver el contenido -->
        <div *ngIf="!esAdmin" class="access-denied" role="alert">
          <span class="access-icon"><i class="bi bi-shield-lock"></i></span>
          <h2>Acceso restringido</h2>
          <p>Solo los administradores pueden gestionar usuarios.</p>
        </div>

        <ng-container *ngIf="esAdmin">
          <!-- RESUMEN -->
          <section class="stats" aria-label="Resumen de usuarios">
            <div class="stat">
              <span class="stat-icon tone-blue"><i class="bi bi-people"></i></span>
              <div class="stat-text">
                <span class="stat-label">Total de usuarios</span>
                <strong class="stat-value">{{ usuarios.length }}</strong>
              </div>
            </div>

            <div class="stat">
              <span class="stat-icon tone-green"><i class="bi bi-person-check"></i></span>
              <div class="stat-text">
                <span class="stat-label">Usuarios activos</span>
                <strong class="stat-value">{{ usuariosActivos }}</strong>
              </div>
            </div>

            <div class="stat">
              <span class="stat-icon tone-rose"><i class="bi bi-shield-lock"></i></span>
              <div class="stat-text">
                <span class="stat-label">Administradores</span>
                <strong class="stat-value">{{ administradores }}</strong>
              </div>
            </div>
          </section>

          <!-- MENSAJES -->
          <div *ngIf="mensajeExito" class="notice notice-success" role="status">
            <i class="bi bi-check-circle-fill notice-icon"></i>
            <span class="notice-text">{{ mensajeExito }}</span>
            <button
              type="button"
              class="icon-btn"
              aria-label="Cerrar mensaje"
              (click)="mensajeExito = ''"
            >
              <i class="bi bi-x-lg"></i>
            </button>
          </div>
          <div *ngIf="mensajeError" class="notice notice-error" role="alert">
            <i class="bi bi-exclamation-triangle-fill notice-icon"></i>
            <span class="notice-text">{{ mensajeError }}</span>
            <button
              type="button"
              class="icon-btn"
              aria-label="Cerrar mensaje"
              (click)="mensajeError = ''"
            >
              <i class="bi bi-x-lg"></i>
            </button>
          </div>

          <!-- TABLA -->
          <section class="panel">
            <!-- FILTROS -->
            <div class="toolbar">
              <div class="search">
                <i class="bi bi-search"></i>
                <input
                  type="search"
                  placeholder="Buscar por nombre o correo"
                  aria-label="Buscar por nombre o correo"
                  [(ngModel)]="busqueda"
                  (input)="filtrarUsuarios()"
                />
              </div>

              <div class="segmented" role="group" aria-label="Filtrar por rol">
                <button
                  type="button"
                  [class.is-active]="filtroRol === 'TODOS'"
                  (click)="filtroRol = 'TODOS'; filtrarUsuarios()"
                >
                  Todos
                </button>
                <button
                  type="button"
                  [class.is-active]="filtroRol === 'ROLE_ADMINISTRADOR'"
                  (click)="filtroRol = 'ROLE_ADMINISTRADOR'; filtrarUsuarios()"
                >
                  Administrador
                </button>
                <button
                  type="button"
                  [class.is-active]="filtroRol === 'ROLE_OPERADOR'"
                  (click)="filtroRol = 'ROLE_OPERADOR'; filtrarUsuarios()"
                >
                  Operador
                </button>
                <button
                  type="button"
                  [class.is-active]="filtroRol === 'ROLE_CLIENTE'"
                  (click)="filtroRol = 'ROLE_CLIENTE'; filtrarUsuarios()"
                >
                  Cliente
                </button>
              </div>
            </div>

            <div *ngIf="cargando" class="loading">
              <div class="spinner-border spinner-border-sm"></div>
              Cargando usuarios de Entra ID...
            </div>

            <div class="table-wrap" *ngIf="!cargando">
              <table class="users-table">
                <thead>
                  <tr>
                    <th>Usuario</th>
                    <th>Rol</th>
                    <th>Estado</th>
                    <th class="col-actions">Acciones</th>
                  </tr>
                </thead>

                <tbody>
                  <tr *ngFor="let usuario of usuariosFiltrados">
                    <td>
                      <div class="user-cell">
                        <div class="user-avatar" [attr.data-rol]="usuario.rol">
                          {{ usuario.nombre.charAt(0).toUpperCase() }}
                        </div>
                        <div class="user-cell-text">
                          <span class="user-cell-name">{{ usuario.nombre }}</span>
                          <span class="user-cell-mail">{{ usuario.correo }}</span>
                        </div>
                      </div>
                    </td>

                    <td>
                      <span
                        *ngIf="usuario.rol === 'ROLE_ADMINISTRADOR'"
                        class="chip chip-admin"
                      >
                        Administrador
                      </span>
                      <span
                        *ngIf="usuario.rol === 'ROLE_OPERADOR'"
                        class="chip chip-operator"
                      >
                        Operador
                      </span>
                      <span
                        *ngIf="usuario.rol === 'ROLE_CLIENTE'"
                        class="chip chip-client"
                      >
                        Cliente
                      </span>
                    </td>

                    <td>
                      <span *ngIf="usuario.estado === 'ACTIVO'" class="status">
                        <span class="dot dot-on"></span>
                        Activo
                      </span>
                      <span *ngIf="usuario.estado === 'INACTIVO'" class="status">
                        <span class="dot dot-off"></span>
                        Inactivo
                      </span>
                    </td>

                    <td class="col-actions">
                      <button
                        class="btn-ghost"
                        title="Editar nombre y rol"
                        (click)="abrirModalEditar(usuario)"
                      >
                        <i class="bi bi-pencil"></i>
                        Editar
                      </button>
                    </td>
                  </tr>

                  <tr *ngIf="usuariosFiltrados.length === 0">
                    <td colspan="4">
                      <div class="empty">
                        <span class="empty-icon"><i class="bi bi-search"></i></span>
                        <strong>No se encontraron usuarios</strong>
                        <span>Prueba con otro nombre o cambia el filtro de rol.</span>
                      </div>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div class="panel-foot" *ngIf="!cargando">
              Mostrando {{ usuariosFiltrados.length }} de {{ usuarios.length }} usuarios
            </div>
          </section>

          <aside class="note">
            <i class="bi bi-info-circle-fill"></i>
            <p>
              <strong>Microsoft Entra ID:</strong>
              los usuarios se gestionan directamente en Entra ID (crear,
              renombrar y rol). El rol se guarda en el campo
              <em>jobTitle</em>. Para que el token JWT incluya el rol de
              aplicación, asígnalo también como
              <strong>App role</strong> en el registro de la API en Azure
              Portal.
            </p>
          </aside>
        </ng-container>
      </div>
    </main>

    <!-- ========== MODAL CREAR ========== -->
    <div class="dialog-layer" *ngIf="mostrarModalCrear" (click)="cerrarModales()">
      <div
        class="dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="titulo-crear"
        (click)="$event.stopPropagation()"
      >
        <div class="dialog-header">
          <div class="dialog-title">
            <span class="dialog-icon"><i class="bi bi-person-plus"></i></span>
            <div>
              <h2 id="titulo-crear">Nuevo usuario</h2>
              <p>Se crea directamente en Entra ID.</p>
            </div>
          </div>
          <button
            type="button"
            class="icon-btn"
            aria-label="Cerrar"
            (click)="cerrarModales()"
          >
            <i class="bi bi-x-lg"></i>
          </button>
        </div>

        <div class="dialog-body">
          <div class="field">
            <label class="field-label" for="crear-nombre">Nombre completo</label>
            <input
              id="crear-nombre"
              type="text"
              class="field-input"
              [(ngModel)]="formCrear.nombre"
              placeholder="Ej: Juan Pérez"
              autocomplete="off"
              required
            />
          </div>

          <div class="field">
            <label class="field-label" for="crear-correo">
              Correo (User Principal Name)
            </label>
            <input
              id="crear-correo"
              type="email"
              class="field-input"
              [(ngModel)]="formCrear.correo"
              placeholder="usuario@tudominio.com"
              autocomplete="off"
              required
            />
            <small class="field-hint">
              Debe ser un dominio verificado en tu tenant.
            </small>
          </div>

          <fieldset class="field role-picker">
            <legend class="field-label">Rol</legend>
            <label class="role-option">
              <input
                type="radio"
                name="rolCrear"
                value="ROLE_CLIENTE"
                [(ngModel)]="formCrear.rol"
              />
              <span class="role-option-body">
                <strong>Cliente</strong>
                <small>Acceso a pedidos</small>
              </span>
            </label>
            <label class="role-option">
              <input
                type="radio"
                name="rolCrear"
                value="ROLE_OPERADOR"
                [(ngModel)]="formCrear.rol"
              />
              <span class="role-option-body">
                <strong>Operador</strong>
                <small>Pedidos y catálogo</small>
              </span>
            </label>
            <label class="role-option">
              <input
                type="radio"
                name="rolCrear"
                value="ROLE_ADMINISTRADOR"
                [(ngModel)]="formCrear.rol"
              />
              <span class="role-option-body">
                <strong>Administrador</strong>
                <small>Todo, incluido usuarios</small>
              </span>
            </label>
          </fieldset>

          <div class="field">
            <label class="field-label" for="crear-password">
              Contraseña temporal
            </label>
            <input
              id="crear-password"
              type="text"
              class="field-input"
              [(ngModel)]="formCrear.passwordTemporal"
              placeholder="Mín. 8 caracteres, mayúscula, número y símbolo"
              autocomplete="off"
              required
            />
            <small class="field-hint">
              El usuario deberá cambiarla en el próximo inicio de sesión.
            </small>
          </div>
        </div>

        <div class="dialog-footer">
          <button type="button" class="btn-soft" (click)="cerrarModales()">
            Cancelar
          </button>
          <button
            type="button"
            class="btn-brand"
            (click)="crearUsuario()"
            [disabled]="guardando || !formCrearValido()"
          >
            <span *ngIf="guardando" class="spinner-border spinner-border-sm"></span>
            Crear usuario
          </button>
        </div>
      </div>
    </div>

    <!-- ========== MODAL EDITAR ========== -->
    <div class="dialog-layer" *ngIf="mostrarModalEditar" (click)="cerrarModales()">
      <div
        class="dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="titulo-editar"
        (click)="$event.stopPropagation()"
      >
        <div class="dialog-header">
          <div class="dialog-title">
            <span class="dialog-icon dialog-icon-avatar">
              {{ (usuarioEditando?.nombre || "U").charAt(0).toUpperCase() }}
            </span>
            <div>
              <h2 id="titulo-editar">Editar usuario</h2>
              <p>Cambia el nombre o el rol de la cuenta.</p>
            </div>
          </div>
          <button
            type="button"
            class="icon-btn"
            aria-label="Cerrar"
            (click)="cerrarModales()"
          >
            <i class="bi bi-x-lg"></i>
          </button>
        </div>

        <div class="dialog-body">
          <div class="field">
            <label class="field-label" for="editar-correo">Correo</label>
            <input
              id="editar-correo"
              type="text"
              class="field-input"
              [value]="usuarioEditando?.correo"
              disabled
            />
          </div>

          <div class="field">
            <label class="field-label" for="editar-nombre">Nombre</label>
            <input
              id="editar-nombre"
              type="text"
              class="field-input"
              [(ngModel)]="formEditar.nombre"
              autocomplete="off"
              required
            />
          </div>

          <fieldset class="field role-picker">
            <legend class="field-label">Rol</legend>
            <label class="role-option">
              <input
                type="radio"
                name="rolEditar"
                value="ROLE_CLIENTE"
                [(ngModel)]="formEditar.rol"
              />
              <span class="role-option-body">
                <strong>Cliente</strong>
                <small>Acceso a pedidos</small>
              </span>
            </label>
            <label class="role-option">
              <input
                type="radio"
                name="rolEditar"
                value="ROLE_OPERADOR"
                [(ngModel)]="formEditar.rol"
              />
              <span class="role-option-body">
                <strong>Operador</strong>
                <small>Pedidos y catálogo</small>
              </span>
            </label>
            <label class="role-option">
              <input
                type="radio"
                name="rolEditar"
                value="ROLE_ADMINISTRADOR"
                [(ngModel)]="formEditar.rol"
              />
              <span class="role-option-body">
                <strong>Administrador</strong>
                <small>Todo, incluido usuarios</small>
              </span>
            </label>
          </fieldset>
        </div>

        <div class="dialog-footer">
          <button type="button" class="btn-soft" (click)="cerrarModales()">
            Cancelar
          </button>
          <button
            type="button"
            class="btn-brand"
            (click)="guardarEdicion()"
            [disabled]="guardando || !formEditar.nombre.trim()"
          >
            <span *ngIf="guardando" class="spinner-border spinner-border-sm"></span>
            Guardar cambios
          </button>
        </div>
      </div>
    </div>
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
        --brand-dark: #166534;
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
        padding: 2rem 1.25rem 3rem;
      }

      .page-head {
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        gap: 1rem;
        margin-bottom: 1.5rem;
      }

      .page-head h1 {
        margin: 0 0 0.3rem;
        font-size: 1.75rem;
        font-weight: 700;
        letter-spacing: -0.02em;
        color: var(--text);
      }

      .page-head p {
        margin: 0;
        max-width: 60ch;
        color: var(--muted);
        font-size: 0.95rem;
        line-height: 1.5;
      }

      /* ---------- Botones ---------- */
      .btn-brand,
      .btn-soft,
      .btn-ghost {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 0.5rem;
        font: inherit;
        font-size: 0.9rem;
        font-weight: 600;
        line-height: 1;
        border-radius: 10px;
        cursor: pointer;
        white-space: nowrap;
        transition: background-color 0.15s ease, border-color 0.15s ease,
          box-shadow 0.15s ease, color 0.15s ease;
      }

      .btn-brand {
        padding: 0.7rem 1.1rem;
        background: var(--brand);
        color: #fff;
        border: 1px solid var(--brand);
        box-shadow: 0 1px 2px rgba(21, 128, 61, 0.35);
      }

      .btn-brand:hover:not(:disabled) {
        background: var(--brand-dark);
        border-color: var(--brand-dark);
      }

      .btn-soft {
        padding: 0.7rem 1.1rem;
        background: #fff;
        color: var(--text);
        border: 1px solid var(--line);
      }

      .btn-soft:hover:not(:disabled) {
        background: #f6f7fb;
        border-color: #cfd2de;
      }

      .btn-ghost {
        padding: 0.45rem 0.8rem;
        font-size: 0.82rem;
        background: transparent;
        color: var(--text);
        border: 1px solid var(--line);
      }

      .btn-ghost:hover {
        background: var(--brand-soft);
        border-color: #b7e0c5;
        color: var(--brand-dark);
      }

      .btn-brand:disabled,
      .btn-soft:disabled {
        opacity: 0.55;
        cursor: not-allowed;
        box-shadow: none;
      }

      .icon-btn {
        flex-shrink: 0;
        width: 32px;
        height: 32px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: 0;
        border-radius: 8px;
        background: transparent;
        color: var(--muted);
        cursor: pointer;
        transition: background-color 0.15s ease, color 0.15s ease;
      }

      .icon-btn:hover {
        background: rgba(15, 23, 42, 0.08);
        color: var(--text);
      }

      /* ---------- Acceso restringido ---------- */
      .access-denied {
        display: flex;
        flex-direction: column;
        align-items: center;
        text-align: center;
        padding: 3.5rem 1.5rem;
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 14px;
      }

      .access-icon {
        width: 56px;
        height: 56px;
        border-radius: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: #fdf1d8;
        color: #8a5a00;
        font-size: 1.6rem;
        margin-bottom: 1rem;
      }

      .access-denied h2 {
        margin: 0 0 0.35rem;
        font-size: 1.2rem;
        font-weight: 700;
      }

      .access-denied p {
        margin: 0;
        color: var(--muted);
      }

      /* ---------- Resumen ---------- */
      .stats {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 14px;
        margin-bottom: 1.25rem;
      }

      .stat {
        display: flex;
        align-items: center;
        gap: 0.9rem;
        padding: 1.1rem 1.35rem;
      }

      .stat + .stat {
        border-left: 1px solid var(--line);
      }

      .stat-icon {
        width: 42px;
        height: 42px;
        flex-shrink: 0;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.2rem;
      }

      .tone-blue {
        background: #e4edff;
        color: #2f55b8;
      }

      .tone-green {
        background: var(--brand-soft);
        color: var(--brand);
      }

      .tone-rose {
        background: #fde8ec;
        color: #b4233f;
      }

      .stat-text {
        display: flex;
        flex-direction: column;
        gap: 0.1rem;
      }

      .stat-label {
        font-size: 0.82rem;
        color: var(--muted);
      }

      .stat-value {
        font-size: 1.55rem;
        font-weight: 700;
        line-height: 1.1;
        letter-spacing: -0.02em;
      }

      /* ---------- Mensajes ---------- */
      .notice {
        display: flex;
        align-items: flex-start;
        gap: 0.7rem;
        padding: 0.8rem 0.85rem 0.8rem 1rem;
        border-radius: 12px;
        border: 1px solid;
        margin-bottom: 1.25rem;
        font-size: 0.92rem;
      }

      .notice-icon {
        margin-top: 0.1rem;
      }

      .notice-text {
        flex: 1;
        line-height: 1.45;
      }

      .notice-success {
        background: #eaf7ef;
        border-color: #b9e3c9;
        color: #14532d;
      }

      .notice-error {
        background: #fdecef;
        border-color: #f5c2cb;
        color: #8c1d33;
      }

      /* ---------- Panel / tabla ---------- */
      .panel {
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 14px;
        overflow: hidden;
      }

      .toolbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 0.75rem;
        padding: 1rem 1.25rem;
        border-bottom: 1px solid var(--line);
      }

      .search {
        position: relative;
        flex: 1 1 260px;
        max-width: 420px;
      }

      .search i {
        position: absolute;
        left: 0.9rem;
        top: 50%;
        transform: translateY(-50%);
        color: var(--muted);
        pointer-events: none;
      }

      .search input {
        width: 100%;
        padding: 0.6rem 0.9rem 0.6rem 2.45rem;
        font: inherit;
        font-size: 0.92rem;
        color: var(--text);
        background: #f7f8fb;
        border: 1px solid var(--line);
        border-radius: 10px;
        transition: border-color 0.15s ease, background-color 0.15s ease;
      }

      .search input::placeholder {
        color: #8a90a5;
      }

      .search input:focus {
        outline: none;
        background: #fff;
        border-color: var(--focus);
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
      }

      .segmented {
        display: inline-flex;
        flex-wrap: wrap;
        gap: 0.25rem;
        padding: 0.25rem;
        background: #f1f2f7;
        border-radius: 10px;
      }

      .segmented button {
        font: inherit;
        font-size: 0.83rem;
        font-weight: 600;
        padding: 0.4rem 0.8rem;
        border: 0;
        border-radius: 8px;
        background: transparent;
        color: var(--muted);
        cursor: pointer;
        transition: background-color 0.15s ease, color 0.15s ease;
      }

      .segmented button:hover {
        color: var(--text);
      }

      .segmented button.is-active {
        background: #fff;
        color: var(--text);
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.12);
      }

      .loading {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 0.6rem;
        padding: 3rem 1rem;
        color: var(--muted);
      }

      .table-wrap {
        overflow-x: auto;
      }

      .users-table {
        width: 100%;
        border-collapse: collapse;
      }

      .users-table th {
        padding: 0.7rem 1.25rem;
        text-align: left;
        font-size: 0.78rem;
        font-weight: 600;
        color: var(--muted);
        background: #fafbfd;
        border-bottom: 1px solid var(--line);
        white-space: nowrap;
      }

      .users-table td {
        padding: 0.85rem 1.25rem;
        border-bottom: 1px solid #eef0f5;
        vertical-align: middle;
      }

      .users-table tbody tr:last-child td {
        border-bottom: 0;
      }

      .users-table tbody tr:hover {
        background: #fafbfe;
      }

      .col-actions {
        text-align: right !important;
      }

      .user-cell {
        display: flex;
        align-items: center;
        gap: 0.8rem;
      }

      .user-avatar {
        width: 40px;
        height: 40px;
        flex-shrink: 0;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 0.95rem;
        background: #e9ebf3;
        color: #4a4f66;
      }

      .user-avatar[data-rol="ROLE_ADMINISTRADOR"] {
        background: #fde8ec;
        color: #b4233f;
      }

      .user-avatar[data-rol="ROLE_OPERADOR"] {
        background: #fdf1d8;
        color: #8a5a00;
      }

      .user-avatar[data-rol="ROLE_CLIENTE"] {
        background: #e4edff;
        color: #2f55b8;
      }

      .user-cell-text {
        display: flex;
        flex-direction: column;
        min-width: 0;
      }

      .user-cell-name {
        font-weight: 600;
        font-size: 0.95rem;
      }

      .user-cell-mail {
        font-size: 0.83rem;
        color: var(--muted);
        overflow-wrap: anywhere;
      }

      .chip {
        display: inline-flex;
        align-items: center;
        padding: 0.25rem 0.7rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 600;
        white-space: nowrap;
      }

      .chip-admin {
        background: #fde8ec;
        color: #b4233f;
      }

      .chip-operator {
        background: #fdf1d8;
        color: #8a5a00;
      }

      .chip-client {
        background: #e4edff;
        color: #2f55b8;
      }

      .status {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        font-size: 0.88rem;
        white-space: nowrap;
      }

      .dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
      }

      .dot-on {
        background: #22c55e;
        box-shadow: 0 0 0 3px rgba(34, 197, 94, 0.18);
      }

      .dot-off {
        background: #a0a5b8;
      }

      .empty {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 0.3rem;
        padding: 2.75rem 1rem;
        text-align: center;
        color: var(--muted);
        font-size: 0.9rem;
      }

      .empty strong {
        color: var(--text);
        font-size: 1rem;
      }

      .empty-icon {
        width: 48px;
        height: 48px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: #f1f2f7;
        color: var(--muted);
        font-size: 1.25rem;
        margin-bottom: 0.5rem;
      }

      .panel-foot {
        padding: 0.8rem 1.25rem;
        border-top: 1px solid var(--line);
        background: #fafbfd;
        font-size: 0.82rem;
        color: var(--muted);
      }

      .note {
        display: flex;
        gap: 0.75rem;
        margin-top: 1.25rem;
        padding: 1rem 1.15rem;
        background: #eef4ff;
        border: 1px solid #d3e2ff;
        border-radius: 12px;
        color: #243b6b;
        font-size: 0.88rem;
        line-height: 1.55;
      }

      .note i {
        margin-top: 0.15rem;
        color: #2f55b8;
      }

      .note p {
        margin: 0;
      }

      /* ---------- Diálogos ---------- */
      .dialog-layer {
        position: fixed;
        inset: 0;
        z-index: 1050;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 1rem;
        background: rgba(20, 21, 36, 0.55);
        backdrop-filter: blur(2px);
        animation: layer-in 0.15s ease-out;
      }

      .dialog {
        width: 100%;
        max-width: 500px;
        max-height: calc(100vh - 2rem);
        display: flex;
        flex-direction: column;
        background: var(--surface);
        border-radius: 16px;
        box-shadow: 0 24px 60px rgba(10, 11, 25, 0.35);
        animation: dialog-in 0.18s ease-out;
      }

      .dialog-header {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 1rem;
        padding: 1.25rem 1.35rem 1rem;
      }

      .dialog-title {
        display: flex;
        align-items: center;
        gap: 0.85rem;
      }

      .dialog-title h2 {
        margin: 0;
        font-size: 1.1rem;
        font-weight: 700;
      }

      .dialog-title p {
        margin: 0.1rem 0 0;
        font-size: 0.85rem;
        color: var(--muted);
      }

      .dialog-icon {
        width: 44px;
        height: 44px;
        flex-shrink: 0;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: var(--brand-soft);
        color: var(--brand);
        font-size: 1.25rem;
      }

      .dialog-icon-avatar {
        background: #e9ebf3;
        color: #4a4f66;
        font-weight: 700;
        font-size: 1.05rem;
      }

      .dialog-body {
        padding: 0.25rem 1.35rem 1.25rem;
        overflow-y: auto;
      }

      .dialog-footer {
        display: flex;
        justify-content: flex-end;
        gap: 0.6rem;
        padding: 1rem 1.35rem;
        border-top: 1px solid var(--line);
        background: #fafbfd;
        border-radius: 0 0 16px 16px;
      }

      .field {
        margin: 0 0 1.1rem;
        padding: 0;
        border: 0;
        min-width: 0;
      }

      .field-label {
        display: block;
        float: none;
        width: auto;
        padding: 0;
        margin: 0 0 0.4rem;
        font-size: 0.85rem;
        font-weight: 600;
        color: var(--text);
      }

      .field-input {
        width: 100%;
        padding: 0.65rem 0.85rem;
        font: inherit;
        font-size: 0.92rem;
        color: var(--text);
        background: #fff;
        border: 1px solid #d3d6e2;
        border-radius: 10px;
        transition: border-color 0.15s ease, box-shadow 0.15s ease;
      }

      .field-input::placeholder {
        color: #8a90a5;
      }

      .field-input:focus {
        outline: none;
        border-color: var(--focus);
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
      }

      .field-input:disabled {
        background: #f3f4f8;
        color: var(--muted);
        cursor: not-allowed;
      }

      .field-hint {
        display: block;
        margin-top: 0.35rem;
        font-size: 0.8rem;
        color: var(--muted);
      }

      .role-picker {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.6rem;
      }

      .role-picker .field-label {
        grid-column: 1 / -1;
        margin-bottom: 0;
      }

      .role-option {
        position: relative;
        display: block;
        cursor: pointer;
      }

      .role-option input {
        position: absolute;
        opacity: 0;
        inset: 0;
        margin: 0;
        cursor: pointer;
      }

      .role-option-body {
        display: flex;
        flex-direction: column;
        gap: 0.2rem;
        height: 100%;
        padding: 0.7rem 0.8rem;
        border: 1px solid #d3d6e2;
        border-radius: 10px;
        background: #fff;
        transition: border-color 0.15s ease, background-color 0.15s ease,
          box-shadow 0.15s ease;
      }

      .role-option-body strong {
        font-size: 0.88rem;
      }

      .role-option-body small {
        font-size: 0.75rem;
        line-height: 1.35;
        color: var(--muted);
      }

      .role-option:hover .role-option-body {
        border-color: #aab0c4;
      }

      .role-option input:checked + .role-option-body {
        border-color: var(--brand);
        background: var(--brand-soft);
        box-shadow: 0 0 0 1px var(--brand);
      }

      .role-option input:focus-visible + .role-option-body {
        outline: 2px solid var(--focus);
        outline-offset: 2px;
      }

      @keyframes layer-in {
        from {
          opacity: 0;
        }
      }

      @keyframes dialog-in {
        from {
          opacity: 0;
          transform: translateY(8px) scale(0.98);
        }
      }

      @media (prefers-reduced-motion: reduce) {
        .dialog-layer,
        .dialog {
          animation: none;
        }
      }

      /* ---------- Responsive ---------- */
      @media (max-width: 900px) {
        .user-text {
          display: none;
        }
      }

      @media (max-width: 768px) {
        .page-head {
          flex-direction: column;
          align-items: stretch;
        }

        .stats {
          grid-template-columns: 1fr;
        }

        .stat + .stat {
          border-left: 0;
          border-top: 1px solid var(--line);
        }

        .search {
          max-width: none;
        }

        .segmented {
          width: 100%;
        }

        .segmented button {
          flex: 1;
        }

        .role-picker {
          grid-template-columns: 1fr;
        }
      }

      @media (max-width: 600px) {
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