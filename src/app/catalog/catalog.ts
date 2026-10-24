import { Component, OnInit, ChangeDetectorRef } from "@angular/core";

import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { FormsModule } from "@angular/forms";

import { AuthService } from "../core/auth.service";
import { MsalService } from "@azure/msal-angular";

import { CatalogoService } from "../services/catalogo-service";
import { Producto } from "../services/catalogo";

@Component({
  selector: "app-catalog",
  standalone: true,
  imports: [
    CommonModule,
    RouterLink,
    FormsModule
  ],

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
            class="nav-item active"
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

        <!-- ========== USUARIO ========== -->
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
          <h1>Catálogo de productos</h1>
          <p *ngIf="esAdmin">Crea productos y mantén el stock al día.</p>
          <p *ngIf="esOperador && !esAdmin">
            Revisa los productos y agrega stock cuando haga falta.
          </p>
        </header>

        <!-- ========== SIN PERMISOS ========== -->
        <div *ngIf="!esAdmin && !esOperador" class="access-denied" role="alert">
          <span class="access-icon"><i class="bi bi-shield-lock"></i></span>
          <h2>Acceso restringido</h2>
          <p>No tienes permiso para administrar el catálogo.</p>
        </div>

        <!-- ========== CREAR PRODUCTO - SOLO ADMIN ========== -->
        <section
          *ngIf="esAdmin"
          class="panel"
          aria-labelledby="titulo-crear-producto"
        >
          <div class="panel-head">
            <div>
              <h2 id="titulo-crear-producto" class="panel-title">
                Crear producto
              </h2>
              <p class="panel-sub">Agrega un nuevo producto al catálogo.</p>
            </div>
          </div>

          <div class="product-form">
            <!-- NOMBRE -->
            <div class="field">
              <label class="field-label" for="producto-nombre">
                Nombre del producto
              </label>
              <input
                id="producto-nombre"
                type="text"
                class="field-input"
                [(ngModel)]="nuevoProducto.nombre"
                name="nombreProducto"
                placeholder="Ej: Producto nuevo"
                autocomplete="off"
              />
            </div>

            <!-- PRECIO -->
            <div class="field">
              <label class="field-label" for="producto-precio">Precio</label>
              <div class="money">
                <span class="money-sign">$</span>
                <input
                  id="producto-precio"
                  type="number"
                  class="field-input"
                  min="0"
                  [(ngModel)]="nuevoProducto.precio"
                  name="precioProducto"
                  placeholder="Ej: 5000"
                />
              </div>
            </div>

            <!-- STOCK INICIAL -->
            <div class="field">
              <label class="field-label" for="producto-stock">
                Stock inicial
              </label>
              <input
                id="producto-stock"
                type="number"
                class="field-input"
                min="0"
                [(ngModel)]="nuevoProducto.stock"
                name="stockProducto"
                placeholder="Ej: 10"
              />
            </div>

            <!-- BOTÓN -->
            <button
              type="button"
              class="btn-brand"
              (click)="crearProducto()"
              [disabled]="
                !nuevoProducto.nombre.trim() ||
                nuevoProducto.precio < 0 ||
                nuevoProducto.stock < 0
              "
            >
              <i class="bi bi-plus-lg"></i>
              Crear producto
            </button>
          </div>

          <p
            class="form-hint is-error"
            *ngIf="nuevoProducto.precio < 0 || nuevoProducto.stock < 0"
          >
            <i class="bi bi-exclamation-circle"></i>
            El precio y el stock no pueden ser negativos.
          </p>
        </section>

        <!-- ========== GESTIÓN DE STOCK ========== -->
        <section
          *ngIf="esAdmin || esOperador"
          class="panel"
          aria-labelledby="titulo-stock"
        >
          <div class="panel-head">
            <div>
              <h2 id="titulo-stock" class="panel-title">Gestión de stock</h2>
              <p class="panel-sub">Agrega unidades al stock de los productos.</p>
            </div>

            <div class="search" *ngIf="productos.length > 0">
              <i class="bi bi-search"></i>
              <input
                type="search"
                placeholder="Buscar por nombre o ID"
                aria-label="Buscar por nombre o ID"
                [(ngModel)]="busqueda"
              />
            </div>
          </div>

          <!-- Resumen por nivel de stock (también filtra la lista) -->
          <div
            class="pipeline"
            role="group"
            aria-label="Filtrar por nivel de stock"
            *ngIf="productos.length > 0"
          >
            <button
              type="button"
              class="stage stage-all"
              [class.is-active]="filtroStock === 'TODOS'"
              (click)="filtroStock = 'TODOS'"
            >
              <span class="stage-count">{{ productos.length }}</span>
              <span class="stage-label">Todos</span>
            </button>

            <button
              *ngFor="let nivel of niveles"
              type="button"
              class="stage"
              [attr.data-nivel]="nivel"
              [class.is-active]="filtroStock === nivel"
              (click)="filtroStock = nivel"
            >
              <span class="stage-count">{{ contarPorNivel(nivel) }}</span>
              <span class="stage-label">
                <span class="dot"></span>
                {{ etiquetaNivel(nivel) }}
              </span>
            </button>
          </div>

          <!-- SIN PRODUCTOS -->
          <div *ngIf="productos.length === 0" class="empty">
            <span class="empty-icon"><i class="bi bi-box-seam"></i></span>
            <strong>No hay productos registrados.</strong>
            <span *ngIf="esAdmin">Crea el primero con el formulario de arriba.</span>
            <span *ngIf="!esAdmin">Cuando un administrador cree productos, aparecerán aquí.</span>
          </div>

          <!-- SIN RESULTADOS CON EL FILTRO -->
          <div
            *ngIf="productos.length > 0 && productosFiltrados.length === 0"
            class="empty"
          >
            <span class="empty-icon"><i class="bi bi-search"></i></span>
            <strong>No hay productos con ese filtro.</strong>
            <span>Prueba con otra búsqueda o cambia el nivel de stock.</span>
            <button
              type="button"
              class="btn-soft"
              (click)="busqueda = ''; filtroStock = 'TODOS'"
            >
              Limpiar filtros
            </button>
          </div>

          <!-- TABLA DE PRODUCTOS -->
          <div *ngIf="productosFiltrados.length > 0" class="table-wrap">
            <table class="catalog-table">
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Producto</th>
                  <th>Precio</th>
                  <th>Stock actual</th>
                  <th class="col-actions">Agregar stock</th>
                </tr>
              </thead>

              <tbody>
                <tr *ngFor="let producto of productosFiltrados">
                  <td class="cell-id">#{{ producto.id }}</td>

                  <td>
                    <div class="product-cell">
                      <span
                        class="product-icon"
                        [attr.data-nivel]="nivelStock(producto.stock)"
                      >
                        <i class="bi bi-box-seam"></i>
                      </span>
                      <strong>{{ producto.nombre }}</strong>
                    </div>
                  </td>

                  <td class="cell-price">$ {{ producto.precio }}</td>

                  <td>
                    <div class="level">
                      <span class="level-count">{{ producto.stock }}</span>
                      <span class="pill" [attr.data-nivel]="nivelStock(producto.stock)">
                        <span class="dot"></span>
                        {{ etiquetaNivel(nivelStock(producto.stock)) }}
                      </span>
                    </div>
                  </td>

                  <td class="col-actions">
                    <div class="stock-adder">
                      <input
                        type="number"
                        class="field-input"
                        min="1"
                        [attr.aria-label]="'Cantidad a agregar a ' + producto.nombre"
                        [(ngModel)]="cantidadesStock[producto.id]"
                        [name]="'stock-' + producto.id"
                      />
                      <button
                        type="button"
                        class="btn-step"
                        [disabled]="
                          !cantidadesStock[producto.id] ||
                          cantidadesStock[producto.id] < 1
                        "
                        (click)="aumentarStock(producto)"
                      >
                        <i class="bi bi-plus-lg"></i>
                        Agregar
                      </button>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="panel-foot" *ngIf="productos.length > 0">
            <span>
              Mostrando {{ productosFiltrados.length }} de {{ productos.length }} productos
            </span>
            <span>Stock bajo: 10 unidades o menos.</span>
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
        display: flex;
        flex-wrap: wrap;
        justify-content: space-between;
        gap: 0.5rem;
        padding: 0.8rem 1.25rem;
        border-top: 1px solid var(--line);
        background: #fafbfd;
        font-size: 0.82rem;
        color: var(--muted);
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

      .btn-step {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        font: inherit;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 0.55rem 0.9rem;
        white-space: nowrap;
        color: var(--brand-dark);
        background: var(--brand-soft);
        border: 1px solid #b7e0c5;
        border-radius: 10px;
        cursor: pointer;
        transition: background-color 0.15s ease, border-color 0.15s ease;
      }

      .btn-step:hover:not(:disabled) {
        background: #d3eedd;
        border-color: #8fcfa5;
      }

      .btn-step:disabled {
        opacity: 0.5;
        cursor: not-allowed;
      }

      /* ---------- Formulario ---------- */
      .product-form {
        display: grid;
        grid-template-columns: minmax(0, 2fr) minmax(0, 1fr) minmax(0, 1fr) auto;
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

      .field-input::placeholder {
        color: #8a90a5;
      }

      .field-input:focus {
        outline: none;
        border-color: var(--focus);
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
      }

      .money {
        position: relative;
      }

      .money-sign {
        position: absolute;
        left: 0.85rem;
        top: 50%;
        transform: translateY(-50%);
        color: var(--muted);
        font-weight: 600;
        pointer-events: none;
      }

      .money .field-input {
        padding-left: 1.8rem;
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

      .form-hint.is-error {
        background: #fdecef;
        border-top-color: #f5c2cb;
        color: #8c1d33;
      }

      /* ---------- Búsqueda ---------- */
      .search {
        position: relative;
        flex: 1 1 240px;
        max-width: 340px;
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

      /* ---------- Niveles de stock ---------- */
      [data-nivel="OK"] {
        --c: #15803d;
        --bg: #e7f6ec;
      }

      [data-nivel="BAJO"] {
        --c: #8a5a00;
        --bg: #fdf1d8;
      }

      [data-nivel="AGOTADO"] {
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
        min-width: 120px;
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

      .catalog-table {
        width: 100%;
        border-collapse: collapse;
      }

      .catalog-table th {
        padding: 0.7rem 1.25rem;
        text-align: left;
        font-size: 0.78rem;
        font-weight: 600;
        white-space: nowrap;
        color: var(--muted);
        background: #fafbfd;
        border-bottom: 1px solid var(--line);
      }

      .catalog-table td {
        padding: 0.8rem 1.25rem;
        text-align: left;
        vertical-align: middle;
        font-size: 0.92rem;
        border-bottom: 1px solid #eef0f5;
      }

      .catalog-table tbody tr:last-child td {
        border-bottom: 0;
      }

      .catalog-table tbody tr:hover {
        background: #fafbfe;
      }

      .catalog-table .col-actions {
        text-align: right;
      }

      .cell-id {
        font-weight: 600;
        color: var(--muted);
        white-space: nowrap;
      }

      .cell-price {
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
      }

      .product-cell {
        display: flex;
        align-items: center;
        gap: 0.8rem;
      }

      .product-icon {
        width: 38px;
        height: 38px;
        flex-shrink: 0;
        border-radius: 11px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.05rem;
        background: var(--bg, #eceef4);
        color: var(--c, #4a4f66);
      }

      .level {
        display: inline-flex;
        align-items: center;
        gap: 0.7rem;
      }

      .level-count {
        min-width: 2.2ch;
        font-size: 1.05rem;
        font-weight: 700;
        font-variant-numeric: tabular-nums;
      }

      .stock-adder {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
      }

      .stock-adder .field-input {
        width: 84px;
        padding: 0.5rem 0.65rem;
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

        .product-form {
          grid-template-columns: 1fr 1fr;
        }

        .product-form .field:first-child {
          grid-column: 1 / -1;
        }

        .product-form .btn-brand {
          grid-column: 1 / -1;
        }
      }

      @media (max-width: 600px) {
        .product-form {
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


export class CatalogComponent implements OnInit {

  nombre = "";

  roles: string[] = [];

  esAdmin = false;

  esOperador = false;


  productos: Producto[] = [];

  cantidadesStock: {
    [id: number]: number
  } = {};


  // ==========================
  // NUEVO PRODUCTO
  // ==========================

  nuevoProducto = {
    nombre: "",
    precio: 0,
    stock: 0
  };


  constructor(
    private auth: AuthService,
    private msal: MsalService,
    private cdr: ChangeDetectorRef,
    private catalogoService: CatalogoService,
  ) {}


  async ngOnInit() {

    // ==========================
    // DATOS DEL USUARIO
    // ==========================

    const account =
      this.auth.getActiveAccount();

    this.nombre =
      account?.name ||
      account?.username ||
      "";


    // ==========================
    // ROLES
    // ==========================

    const roles =
      await this.auth.getRoles();

    this.roles =
      roles;


    this.esAdmin =
      roles.includes(
        "ROLE_ADMINISTRADOR"
      );


    this.esOperador =
      roles.includes(
        "ROLE_OPERADOR"
      );


    console.log(
      "Roles en Catálogo:",
      roles
    );

    console.log(
      "Acceso Admin:",
      this.esAdmin
    );

    console.log(
      "Acceso Operador:",
      this.esOperador
    );


    // ==========================
    // CARGAR PRODUCTOS
    // ==========================

    if (
      this.esAdmin ||
      this.esOperador
    ) {

      this.catalogoService
        .listarProductos()
        .subscribe({

          next: (
            productos: Producto[]
          ) => {

            this.productos =
              productos;


            // Cantidad inicial
            // para cada producto

            productos.forEach(
              producto => {

                this.cantidadesStock[
                  producto.id
                ] = 1;

              }
            );


            console.log(
              "Productos cargados:",
              this.productos
            );


            this.cdr.detectChanges();

          },


          error: (error: any) => {

            console.error(
              "❌ Error obteniendo productos:",
              error
            );

          },

        });

    }


    this.cdr.detectChanges();

  }


  // ==========================
  // CREAR PRODUCTO
  // ==========================

  crearProducto() {

    if (
      !this.nuevoProducto.nombre.trim()
    ) {

      console.warn(
        "El nombre del producto es obligatorio."
      );

      return;

    }


    if (
      this.nuevoProducto.precio < 0 ||
      this.nuevoProducto.stock < 0
    ) {

      console.warn(
        "El precio y el stock no pueden ser negativos."
      );

      return;

    }


    console.log(
      "🆕 Creando producto:",
      this.nuevoProducto
    );


    this.catalogoService
      .crearProducto(
        this.nuevoProducto
      )
      .subscribe({

        next: (
          producto: Producto
        ) => {

          console.log(
            "✅ Producto creado:",
            producto
          );


          // Agregar el nuevo producto
          // directamente a la tabla

          this.productos.push(
            producto
          );


          // Cantidad por defecto
          // para agregar stock

          this.cantidadesStock[
            producto.id
          ] = 1;


          // Limpiar formulario

          this.nuevoProducto = {
            nombre: "",
            precio: 0,
            stock: 0
          };


          this.cdr.detectChanges();

        },


        error: (error: any) => {

          console.error(
            "❌ Error creando producto:",
            error
          );

        },

      });

  }


  // ==========================
  // AUMENTAR STOCK
  // ==========================

  aumentarStock(
    producto: Producto
  ) {

    const cantidad =
      this.cantidadesStock[
        producto.id
      ];


    if (
      !cantidad ||
      cantidad <= 0
    ) {

      console.warn(
        "La cantidad debe ser mayor que cero."
      );

      return;

    }


    console.log(
      `📦 Agregando ${cantidad} unidades al producto #${producto.id}`
    );


    this.catalogoService
      .aumentarStock(
        producto.id,
        cantidad
      )
      .subscribe({

        next: (
          productoActualizado: Producto
        ) => {

          console.log(
            "✅ Stock actualizado:",
            productoActualizado
          );


          const indice =
            this.productos.findIndex(
              p =>
                p.id === productoActualizado.id
            );


          if (
            indice !== -1
          ) {

            this.productos[indice] =
              productoActualizado;

          }


          // Reiniciar cantidad

          this.cantidadesStock[
            producto.id
          ] = 1;


          this.cdr.detectChanges();

        },


        error: (error: any) => {

          console.error(
            "❌ Error aumentando stock:",
            error
          );

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
  // (búsqueda, filtro por nivel de stock y resumen).
  // Todo lo anterior queda igual.
  // ======================================================

  busqueda = "";

  filtroStock = "TODOS";

  readonly niveles: string[] = ["OK", "BAJO", "AGOTADO"];

  private readonly etiquetasNivel: Record<string, string> = {
    OK: "Disponible",
    BAJO: "Stock bajo",
    AGOTADO: "Agotado",
  };

  // Mismos umbrales que usaban las insignias originales:
  // 0 = agotado, 1 a 10 = bajo, más de 10 = disponible.
  nivelStock(stock: number): string {
    if (stock <= 0) return "AGOTADO";
    if (stock <= 10) return "BAJO";
    return "OK";
  }

  etiquetaNivel(nivel: string): string {
    return this.etiquetasNivel[nivel] ?? nivel;
  }

  contarPorNivel(nivel: string): number {
    return this.productos.filter((p) => this.nivelStock(p.stock) === nivel)
      .length;
  }

  get productosFiltrados(): Producto[] {
    const texto = this.busqueda.toLowerCase().trim();

    return this.productos.filter((producto) => {
      const coincideNivel =
        this.filtroStock === "TODOS" ||
        this.nivelStock(producto.stock) === this.filtroStock;

      const coincideTexto =
        !texto ||
        String(producto.nombre).toLowerCase().includes(texto) ||
        String(producto.id).toLowerCase().includes(texto);

      return coincideNivel && coincideTexto;
    });
  }

}