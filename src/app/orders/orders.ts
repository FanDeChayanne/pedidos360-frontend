import { Component, OnInit, ChangeDetectorRef } from "@angular/core";
import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { FormsModule } from "@angular/forms";

import { CatalogoService } from "../services/catalogo-service";
import { Producto } from "../services/catalogo";
import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";

import { PedidoService } from "../services/pedido-service";
import { Pedido, CrearPedido } from "../services/pedido";

@Component({
  selector: "app-orders",
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

          <a class="nav-item active" routerLink="/orders">
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

          <a class="nav-item" *ngIf="esAdmin" routerLink="/usuarios">
            <i class="bi bi-people"></i>
            <span>Usuarios</span>
          </a>

          <a class="nav-item" href="javascript:void(0)">
            <i class="bi bi-box"></i>
            <span>TEST</span>
          </a>
        </div>

        <!-- Usuario -->
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
          <h1>Pedidos</h1>
          <p *ngIf="esOperador || esAdmin">
            Aquí se pueden gestionar los pedidos y sus estados.
          </p>
          <p *ngIf="!(esOperador || esAdmin)">
            Crea pedidos y sigue su avance hasta la entrega.
          </p>
        </header>

        <!-- ========== CREAR PEDIDO ========== -->
        <section
          *ngIf="esCliente || esAdmin"
          class="panel form-panel"
          aria-labelledby="titulo-crear"
        >
          <div class="panel-head">
            <div>
              <h2 id="titulo-crear" class="panel-title">Crear pedido</h2>
              <p class="panel-sub">Ingresa los datos del nuevo pedido.</p>
            </div>
          </div>

          <form class="order-form" (ngSubmit)="crearPedido()">
            <!-- Cliente (tomado automáticamente de la cuenta logueada) -->
            <div class="field">
              <label class="field-label" for="pedido-cliente">Cliente</label>
              <input
                id="pedido-cliente"
                type="text"
                class="field-input"
                [value]="nuevoPedido.clienteId"
                name="clienteId"
                readonly
                disabled
              />
            </div>

            <!-- Producto -->
            <div class="field">
              <label class="field-label" for="pedido-producto">Producto</label>
              <select
                id="pedido-producto"
                class="field-input"
                [(ngModel)]="nuevoPedido.productoId"
                name="productoId"
                required
              >
                <option
                  *ngFor="let producto of productos"
                  [ngValue]="producto.id"
                >
                  {{ producto.nombre }} - Precio: {{ producto.precio }} -
                  Stock: {{ producto.stock }}
                </option>
              </select>
            </div>

            <!-- Cantidad -->
            <div class="field">
              <label class="field-label" for="pedido-cantidad">Cantidad</label>
              <input
                id="pedido-cantidad"
                type="number"
                class="field-input"
                [(ngModel)]="nuevoPedido.cantidad"
                name="cantidad"
                min="1"
                required
              />
            </div>

            <!-- Botón -->
            <button
              type="submit"
              class="btn-brand"
              [disabled]="
                productos.length === 0 ||
                !nuevoPedido.cantidad ||
                nuevoPedido.cantidad < 1
              "
            >
              <i class="bi bi-plus-lg"></i>
              Crear pedido
            </button>
          </form>

          <p class="form-hint" *ngIf="productoSeleccionado as producto">
            <i class="bi bi-box-seam"></i>
            {{ producto.nombre }}: precio {{ producto.precio }}, stock disponible
            {{ producto.stock }}.
          </p>
          <p class="form-hint" *ngIf="productos.length === 0">
            <i class="bi bi-info-circle"></i>
            Todavía no hay productos en el catálogo para crear un pedido.
          </p>
        </section>

        <!-- ========== LISTA DE PEDIDOS ========== -->
        <section class="panel" aria-labelledby="titulo-lista">
          <div class="panel-head">
            <h2 id="titulo-lista" class="panel-title">Lista de pedidos</h2>

            <div class="search" *ngIf="pedidos.length > 0">
              <i class="bi bi-search"></i>
              <input
                type="search"
                placeholder="Buscar por ID, cliente o producto"
                aria-label="Buscar por ID, cliente o producto"
                [(ngModel)]="busqueda"
              />
            </div>
          </div>

          <!-- Resumen por estado (también filtra la lista) -->
          <div
            class="pipeline"
            role="group"
            aria-label="Filtrar por estado"
            *ngIf="pedidos.length > 0"
          >
            <button
              type="button"
              class="stage stage-all"
              [class.is-active]="filtroEstado === 'TODOS'"
              (click)="filtroEstado = 'TODOS'"
            >
              <span class="stage-count">{{ pedidos.length }}</span>
              <span class="stage-label">Todos</span>
            </button>

            <button
              *ngFor="let estado of estados"
              type="button"
              class="stage"
              [attr.data-estado]="estado"
              [class.is-active]="filtroEstado === estado"
              (click)="filtroEstado = estado"
            >
              <span class="stage-count">{{ contarPorEstado(estado) }}</span>
              <span class="stage-label">
                <span class="dot"></span>
                {{ etiquetaEstado(estado) }}
              </span>
            </button>
          </div>

          <!-- Sin pedidos -->
          <div *ngIf="pedidos.length === 0" class="empty">
            <span class="empty-icon"><i class="bi bi-cart3"></i></span>
            <strong>No hay pedidos registrados.</strong>
            <span>Los pedidos nuevos aparecerán aquí.</span>
          </div>

          <!-- Sin resultados con el filtro -->
          <div
            *ngIf="pedidos.length > 0 && pedidosFiltrados.length === 0"
            class="empty"
          >
            <span class="empty-icon"><i class="bi bi-search"></i></span>
            <strong>No hay pedidos con ese filtro.</strong>
            <span>Prueba con otra búsqueda o cambia el estado.</span>
            <button
              type="button"
              class="btn-soft"
              (click)="busqueda = ''; filtroEstado = 'TODOS'"
            >
              Limpiar filtros
            </button>
          </div>

          <!-- Tabla -->
          <div *ngIf="pedidosFiltrados.length > 0" class="table-wrap">
            <table class="orders-table">
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Cliente</th>
                  <th>Producto</th>
                  <th class="col-center">Cantidad</th>
                  <th>Estado</th>
                  <th class="col-actions">Acciones</th>
                </tr>
              </thead>

              <tbody>
                <tr *ngFor="let pedido of pedidosFiltrados">
                  <td class="cell-id">#{{ pedido.id }}</td>

                  <td>{{ pedido.clienteId }}</td>

                  <td class="cell-product">
                    {{ obtenerNombreProducto(pedido.productoId) }}
                  </td>

                  <td class="col-center">{{ pedido.cantidad }}</td>

                  <td>
                    <span class="pill" [attr.data-estado]="pedido.estado">
                      <span class="dot"></span>
                      {{ etiquetaEstado(pedido.estado) }}
                    </span>
                  </td>

                  <!-- ACCIONES -->
                  <td class="col-actions">
                    <div
                      *ngIf="
                        (esAdmin || esOperador) &&
                        obtenerSiguienteEstados(pedido.estado).length > 0
                      "
                      class="actions"
                    >
                      <button
                        *ngFor="
                          let siguienteEstado of obtenerSiguienteEstados(
                            pedido.estado
                          )
                        "
                        type="button"
                        class="btn-step"
                        [class.is-danger]="siguienteEstado === 'CANCELADO'"
                        [title]="'Cambiar a ' + etiquetaEstado(siguienteEstado)"
                        (click)="cambiarEstado(pedido, siguienteEstado)"
                      >
                        {{ accionEstado(siguienteEstado) }}
                      </button>
                    </div>

                    <span
                      *ngIf="
                        !(esAdmin || esOperador) ||
                        obtenerSiguienteEstados(pedido.estado).length === 0
                      "
                      class="no-actions"
                    >
                      Sin acciones disponibles
                    </span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="panel-foot" *ngIf="pedidos.length > 0">
            Mostrando {{ pedidosFiltrados.length }} de {{ pedidos.length }} pedidos
          </div>
        </section>
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
        margin-bottom: 1.5rem;
      }

      .page-head h1 {
        margin: 0 0 0.3rem;
        font-size: 1.75rem;
        font-weight: 700;
        letter-spacing: -0.02em;
      }

      .page-head p {
        margin: 0;
        max-width: 60ch;
        color: var(--muted);
        font-size: 0.95rem;
        line-height: 1.5;
      }

      .panel {
        background: var(--surface);
        border: 1px solid var(--line);
        border-radius: 14px;
        overflow: hidden;
        margin-bottom: 1.25rem;
      }

      .panel-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 0.75rem;
        padding: 1.1rem 1.25rem;
      }

      .panel-title {
        margin: 0;
        font-size: 1.05rem;
        font-weight: 700;
      }

      .panel-sub {
        margin: 0.15rem 0 0;
        font-size: 0.87rem;
        color: var(--muted);
      }

      .panel-foot {
        padding: 0.8rem 1.25rem;
        border-top: 1px solid var(--line);
        background: #fafbfd;
        font-size: 0.82rem;
        color: var(--muted);
      }

      /* ---------- Botones ---------- */
      .btn-brand,
      .btn-soft {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 0.5rem;
        font: inherit;
        font-size: 0.9rem;
        font-weight: 600;
        line-height: 1;
        padding: 0.72rem 1.15rem;
        border-radius: 10px;
        cursor: pointer;
        white-space: nowrap;
        transition: background-color 0.15s ease, border-color 0.15s ease;
      }

      .btn-brand {
        background: var(--brand);
        color: #fff;
        border: 1px solid var(--brand);
        box-shadow: 0 1px 2px rgba(21, 128, 61, 0.35);
      }

      .btn-brand:hover:not(:disabled) {
        background: var(--brand-dark);
        border-color: var(--brand-dark);
      }

      .btn-brand:disabled {
        opacity: 0.55;
        cursor: not-allowed;
        box-shadow: none;
      }

      .btn-soft {
        background: #fff;
        color: var(--text);
        border: 1px solid var(--line);
      }

      .btn-soft:hover {
        background: #f6f7fb;
        border-color: #cfd2de;
      }

      /* ---------- Formulario ---------- */
      .order-form {
        display: grid;
        grid-template-columns: minmax(0, 1.3fr) minmax(0, 2fr) minmax(0, 0.8fr) auto;
        gap: 1rem;
        align-items: end;
        padding: 0 1.25rem 1.25rem;
      }

      .field {
        min-width: 0;
      }

      .field-label {
        display: block;
        margin: 0 0 0.4rem;
        font-size: 0.85rem;
        font-weight: 600;
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

      .form-hint {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        margin: 0;
        padding: 0.75rem 1.25rem;
        border-top: 1px solid var(--line);
        background: #fafbfd;
        font-size: 0.85rem;
        color: var(--muted);
      }

      .form-hint i {
        color: var(--brand);
      }

      /* ---------- Búsqueda ---------- */
      .search {
        position: relative;
        flex: 1 1 260px;
        max-width: 380px;
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

      /* ---------- Estados ---------- */
      [data-estado="CREADO"] {
        --c: #2f55b8;
        --bg: #e4edff;
      }

      [data-estado="ACEPTADO"] {
        --c: #0e7490;
        --bg: #dff4f8;
      }

      [data-estado="EN_PREPARACION"] {
        --c: #8a5a00;
        --bg: #fdf1d8;
      }

      [data-estado="DESPACHADO"] {
        --c: #5b3fb5;
        --bg: #eee8fd;
      }

      [data-estado="ENTREGADO"] {
        --c: #15803d;
        --bg: #e7f6ec;
      }

      [data-estado="CANCELADO"] {
        --c: #b4233f;
        --bg: #fde8ec;
      }

      .dot {
        width: 8px;
        height: 8px;
        flex-shrink: 0;
        border-radius: 50%;
        background: var(--c, #8a90a5);
      }

      .pipeline {
        display: flex;
        gap: 0.6rem;
        padding: 0.9rem 1.25rem;
        overflow-x: auto;
        border-top: 1px solid var(--line);
        border-bottom: 1px solid var(--line);
        background: #fafbfd;
      }

      .stage {
        flex: 1 0 auto;
        min-width: 108px;
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        gap: 0.2rem;
        padding: 0.6rem 0.85rem;
        font: inherit;
        text-align: left;
        color: var(--text);
        background: #fff;
        border: 1px solid var(--line);
        border-radius: 12px;
        cursor: pointer;
        transition: border-color 0.15s ease, box-shadow 0.15s ease;
      }

      .stage:hover {
        border-color: #b8bdd0;
      }

      .stage.is-active {
        border-color: var(--c, var(--ink));
        box-shadow: 0 0 0 1px var(--c, var(--ink));
      }

      .stage-all {
        --c: var(--ink);
      }

      .stage[data-estado="CANCELADO"] {
        margin-left: 0.4rem;
      }

      .stage-count {
        font-size: 1.4rem;
        font-weight: 700;
        line-height: 1.1;
        letter-spacing: -0.02em;
      }

      .stage-label {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        font-size: 0.8rem;
        font-weight: 600;
        color: var(--muted);
        white-space: nowrap;
      }

      .pill {
        display: inline-flex;
        align-items: center;
        gap: 0.45rem;
        padding: 0.25rem 0.75rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 600;
        white-space: nowrap;
        background: var(--bg, #eceef4);
        color: var(--c, #4a4f66);
      }

      .pill .dot {
        background: currentColor;
      }

      /* ---------- Tabla ---------- */
      .table-wrap {
        overflow-x: auto;
      }

      .orders-table {
        width: 100%;
        border-collapse: collapse;
      }

      .orders-table th {
        padding: 0.7rem 1.25rem;
        text-align: left;
        font-size: 0.78rem;
        font-weight: 600;
        white-space: nowrap;
        color: var(--muted);
        background: #fafbfd;
        border-bottom: 1px solid var(--line);
      }

      .orders-table td {
        padding: 0.85rem 1.25rem;
        text-align: left;
        vertical-align: middle;
        font-size: 0.92rem;
        border-bottom: 1px solid #eef0f5;
      }

      .orders-table tbody tr:last-child td {
        border-bottom: 0;
      }

      .orders-table tbody tr:hover {
        background: #fafbfe;
      }

      .orders-table .col-center {
        text-align: center;
      }

      .orders-table .col-actions {
        text-align: right;
      }

      .cell-id {
        font-weight: 600;
        color: var(--muted);
        white-space: nowrap;
      }

      .cell-product {
        font-weight: 600;
      }

      .actions {
        display: inline-flex;
        flex-wrap: wrap;
        justify-content: flex-end;
        gap: 0.4rem;
      }

      .btn-step {
        font: inherit;
        font-size: 0.8rem;
        font-weight: 600;
        padding: 0.4rem 0.8rem;
        white-space: nowrap;
        color: var(--brand-dark);
        background: var(--brand-soft);
        border: 1px solid #b7e0c5;
        border-radius: 8px;
        cursor: pointer;
        transition: background-color 0.15s ease, border-color 0.15s ease;
      }

      .btn-step:hover {
        background: #d3eedd;
        border-color: #8fcfa5;
      }

      .btn-step.is-danger {
        color: #b4233f;
        background: #fff;
        border-color: #f1c1ca;
      }

      .btn-step.is-danger:hover {
        background: #fde8ec;
        border-color: #e79aa8;
      }

      .no-actions {
        font-size: 0.82rem;
        color: #8a90a5;
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

      .empty .btn-soft {
        margin-top: 0.7rem;
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

      /* ---------- Responsive ---------- */
      @media (max-width: 900px) {
        .user-text {
          display: none;
        }

        .order-form {
          grid-template-columns: 1fr 1fr;
        }

        .order-form .btn-brand {
          grid-column: 1 / -1;
        }
      }

      @media (max-width: 600px) {
        .order-form {
          grid-template-columns: 1fr;
        }

        .search {
          max-width: none;
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
export class OrdersComponent implements OnInit {
  nombre = "";
  roles: string[] = [];

  esAdmin = false;
  esOperador = false;
  esCliente = false;

  productos: Producto[] = [];
  pedidos: Pedido[] = [];

  nuevoPedido: CrearPedido = {
    clienteId: "",
    productoId: 0,
    cantidad: 1,
  };

  constructor(
    private auth: AuthService,
    private msal: MsalService,
    private cdr: ChangeDetectorRef,
    private pedidoService: PedidoService,
    private catalogoService: CatalogoService,
  ) {}

  async ngOnInit() {
    // ==========================
    // USUARIO
    // ==========================

    const account = this.auth.getActiveAccount();

    this.nombre = account?.name || account?.username || "";

    // El cliente se toma automáticamente del nombre de la cuenta logueada
    // (sin importar el rol: admin, operador o cliente)
    this.nuevoPedido.clienteId = this.nombre;

    // ==========================
    // ROLES
    // ==========================

    const roles = await this.auth.getRoles();

    this.roles = roles;

    this.esAdmin = roles.includes("ROLE_ADMINISTRADOR");

    this.esOperador = roles.includes("ROLE_OPERADOR");

    this.esCliente = roles.includes("ROLE_CLIENTE");

    console.log("Roles en Pedidos:", roles);

    console.log("Acceso Admin:", this.esAdmin);

    console.log("Acceso Operador:", this.esOperador);

    console.log("Acceso Cliente:", this.esCliente);

    // ==========================
    // CARGAR PEDIDOS
    // ==========================

    this.pedidoService.listarPedidos().subscribe({
      next: (pedidos: Pedido[]) => {
        this.pedidos = pedidos;

        console.log("Pedidos cargados:", this.pedidos);

        this.cdr.detectChanges();
      },

      error: (error: any) => {
        console.error("Error obteniendo pedidos:", error);
      },
    });

    // ==========================
    // CARGAR PRODUCTOS
    // ==========================

    this.catalogoService.listarProductos().subscribe({
      next: (productos: Producto[]) => {
        this.productos = productos;

        console.log("Productos cargados:", this.productos);

        // Seleccionar automáticamente
        // el primer producto

        if (this.productos.length > 0) {
          this.nuevoPedido.productoId = this.productos[0].id;
        }

        this.cdr.detectChanges();
      },

      error: (error: any) => {
        console.error("Error obteniendo productos:", error);
      },
    });

    this.cdr.detectChanges();
  }

  // ==========================
  // OBTENER NOMBRE PRODUCTO
  // ==========================

  obtenerNombreProducto(productoId: number): string {
    const producto = this.productos.find((p) => p.id === productoId);

    return producto ? producto.nombre : `Producto #${productoId}`;
  }

  // ==========================
  // CREAR PEDIDO
  // ==========================

  crearPedido() {
    console.log("🚀 Creando pedido:", this.nuevoPedido);

    this.pedidoService.crearPedido(this.nuevoPedido).subscribe({
      next: (pedido: Pedido) => {
        console.log("✅ Pedido creado:", pedido);

        // Agregar el pedido
        // inmediatamente a la tabla

        this.pedidos.push(pedido);

        // Reiniciar formulario (mantener el cliente de la cuenta)
        this.nuevoPedido = {
          clienteId: this.nombre,

          productoId: this.productos.length > 0 ? this.productos[0].id : 0,

          cantidad: 1,
        };

        this.cdr.detectChanges();
      },

      error: (error: any) => {
        console.error("❌ Error creando pedido:", error);
      },
    });
  }

  // ==========================
  // OBTENER SIGUIENTES ESTADOS
  // ==========================

  obtenerSiguienteEstados(estado: string): string[] {
    switch (estado) {
      case "CREADO":
        return ["ACEPTADO", "CANCELADO"];

      case "ACEPTADO":
        return ["EN_PREPARACION", "CANCELADO"];

      case "EN_PREPARACION":
        return ["DESPACHADO", "CANCELADO"];

      case "DESPACHADO":
        return ["ENTREGADO"];

      default:
        return [];
    }
  }

  // ==========================
  // CAMBIAR ESTADO
  // ==========================

  cambiarEstado(pedido: Pedido, nuevoEstado: string) {
    console.log(
      `🔄 Cambiando pedido #${pedido.id}: ${pedido.estado} → ${nuevoEstado}`,
    );

    this.pedidoService.cambiarEstado(pedido.id, nuevoEstado).subscribe({
      next: (pedidoActualizado: Pedido) => {
        console.log("✅ Pedido actualizado:", pedidoActualizado);

        const indice = this.pedidos.findIndex(
          (p) => p.id === pedidoActualizado.id,
        );

        if (indice !== -1) {
          this.pedidos[indice] = pedidoActualizado;
        }

        this.cdr.detectChanges();
      },

      error: (error: any) => {
        console.error("❌ Error cambiando estado del pedido:", error);
      },
    });
  }

  // ==========================
  // LOGOUT
  // ==========================

  logout() {
    this.msal.logoutRedirect();
  }

  // ======================================================
  // AÑADIDO PARA LA NUEVA INTERFAZ
  // (búsqueda, filtro por estado, resumen y textos).
  // Todo lo anterior queda igual.
  // ======================================================

  busqueda = "";
  filtroEstado = "TODOS";

  readonly estados: string[] = [
    "CREADO",
    "ACEPTADO",
    "EN_PREPARACION",
    "DESPACHADO",
    "ENTREGADO",
    "CANCELADO",
  ];

  private readonly etiquetas: Record<string, string> = {
    CREADO: "Creado",
    ACEPTADO: "Aceptado",
    EN_PREPARACION: "En preparación",
    DESPACHADO: "Despachado",
    ENTREGADO: "Entregado",
    CANCELADO: "Cancelado",
  };

  private readonly acciones: Record<string, string> = {
    ACEPTADO: "Aceptar",
    EN_PREPARACION: "Preparar",
    DESPACHADO: "Despachar",
    ENTREGADO: "Marcar entregado",
    CANCELADO: "Cancelar",
  };

  get pedidosFiltrados(): Pedido[] {
    const texto = this.busqueda.toLowerCase().trim();

    return this.pedidos.filter((pedido) => {
      const coincideEstado =
        this.filtroEstado === "TODOS" || pedido.estado === this.filtroEstado;

      const coincideTexto =
        !texto ||
        String(pedido.id).toLowerCase().includes(texto) ||
        String(pedido.clienteId).toLowerCase().includes(texto) ||
        this.obtenerNombreProducto(pedido.productoId)
          .toLowerCase()
          .includes(texto);

      return coincideEstado && coincideTexto;
    });
  }

  get productoSeleccionado(): Producto | undefined {
    return this.productos.find((p) => p.id === this.nuevoPedido.productoId);
  }

  contarPorEstado(estado: string): number {
    return this.pedidos.filter((p) => p.estado === estado).length;
  }

  etiquetaEstado(estado: string): string {
    return this.etiquetas[estado] ?? estado;
  }

  accionEstado(estado: string): string {
    return this.acciones[estado] ?? estado;
  }
}