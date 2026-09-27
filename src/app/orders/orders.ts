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

          <a class="nav-item" *ngIf="esAdmin" routerLink="/usuarios">
            <i class="bi bi-people"></i>
            Usuarios
          </a>

          <a class="nav-item" href="javascript:void(0)">
            <i class="bi bi-box"></i>
            TEST
          </a>
        </div>

        <!-- Usuario -->

        <div class="nav-right">
          <div class="user-info">
            <div class="avatar">
              {{ nombre ? nombre.charAt(0).toUpperCase() : "U" }}
            </div>

            <div class="user-text">
              <div class="user-name">
                {{ nombre || "Usuario" }}
              </div>

              <div class="user-role">
                {{ roles.join(", ") || "Sin roles" }}
              </div>
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
        <h2 class="mb-4">Pedidos</h2>

        <!-- ========== CREAR PEDIDO ========== -->

        <div *ngIf="esCliente || esAdmin" class="card border-0 shadow-sm mb-4">
          <div class="card-body">
            <h5 class="card-title">Crear pedido</h5>

            <p class="card-text text-muted">
              Ingresa los datos del nuevo pedido.
            </p>

            <form class="row g-3" (ngSubmit)="crearPedido()">
              <!-- Cliente (tomado automáticamente de la cuenta logueada) -->

              <div class="col-md-5">
                <label class="form-label"> Cliente </label>

                <input
                  type="text"
                  class="form-control"
                  [value]="nuevoPedido.clienteId"
                  name="clienteId"
                  readonly
                  disabled
                />
                
              </div>

              <!-- Producto -->

              <div class="col-md-3">
                <label class="form-label"> Producto </label>

                <select
                  class="form-select"
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

              <div class="col-md-2">
                <label class="form-label"> Cantidad </label>

                <input
                  type="number"
                  class="form-control"
                  [(ngModel)]="nuevoPedido.cantidad"
                  name="cantidad"
                  min="1"
                  required
                />
              </div>

              <!-- Botón -->

              <div class="col-md-2 d-flex align-items-end">
                <button type="submit" class="btn btn-primary w-100">
                  Crear pedido
                </button>
              </div>
            </form>
          </div>
        </div>

        <!-- ========== INFORMACIÓN OPERADOR / ADMIN ========== -->

        <div *ngIf="esOperador || esAdmin" class="card border-0 shadow-sm mb-4">
          <div class="card-body">
            <h5 class="card-title"></h5>

            <p class="card-text">
              Aquí se pueden gestionar los pedidos y sus estados.
            </p>

            <ul class="mb-0"></ul>
          </div>
        </div>

        <!-- ========== LISTA DE PEDIDOS ========== -->

        <div class="card border-0 shadow-sm">
          <div class="card-body">
            <div class="d-flex justify-content-between align-items-center mb-3">
              <h5 class="card-title mb-0">Lista de pedidos</h5>

              <span class="badge bg-secondary">
                {{ pedidos.length }} pedidos
              </span>
            </div>

            <!-- Sin pedidos -->

            <div *ngIf="pedidos.length === 0" class="text-muted">
              No hay pedidos registrados.
            </div>

            <!-- Tabla -->

            <div *ngIf="pedidos.length > 0" class="table-responsive">
              <table
                class="table table-hover align-middle mb-0"
                style="width: 100%;"
              >
                <thead>
                  <tr>
                    <th>ID</th>

                    <th>Cliente</th>

                    <th>Producto</th>

                    <th class="text-center">Cantidad</th>

                    <th class="text-center">Estado</th>

                    <th>Acciones</th>
                  </tr>
                </thead>

                <tbody>
                  <tr *ngFor="let pedido of pedidos">
                    <td>
                      {{ pedido.id }}
                    </td>

                    <td>
                      {{ pedido.clienteId }}
                    </td>

                    <td>
                      {{ obtenerNombreProducto(pedido.productoId) }}
                    </td>

                    <td class="text-center">
                      {{ pedido.cantidad }}
                    </td>

                    <td class="text-center">
                      <span
                        class="badge"
                        [ngClass]="{
                          'bg-primary': pedido.estado === 'CREADO',
                          'bg-info text-dark': pedido.estado === 'ACEPTADO',
                          'bg-warning text-dark':
                            pedido.estado === 'EN_PREPARACION',
                          'bg-secondary': pedido.estado === 'DESPACHADO',
                          'bg-success': pedido.estado === 'ENTREGADO',
                          'bg-danger': pedido.estado === 'CANCELADO',
                        }"
                      >
                        {{ pedido.estado }}
                      </span>
                    </td>

                    <!-- ACCIONES -->

                    <td>
                      <div
                        *ngIf="
                          (esAdmin || esOperador) &&
                          obtenerSiguienteEstados(pedido.estado).length > 0
                        "
                        class="d-flex flex-wrap gap-2"
                      >
                        <button
                          *ngFor="
                            let siguienteEstado of obtenerSiguienteEstados(
                              pedido.estado
                            )
                          "
                          type="button"
                          class="btn btn-sm"
                          [ngClass]="{
                            'btn-success': siguienteEstado === 'ACEPTADO',
                            'btn-warning': siguienteEstado === 'EN_PREPARACION',
                            'btn-primary': siguienteEstado === 'DESPACHADO',
                            'btn-info': siguienteEstado === 'ENTREGADO',
                            'btn-danger': siguienteEstado === 'CANCELADO',
                          }"
                          (click)="cambiarEstado(pedido, siguienteEstado)"
                        >
                          {{ siguienteEstado }}
                        </button>
                      </div>

                      <span
                        *ngIf="
                          obtenerSiguienteEstados(pedido.estado).length === 0
                        "
                        class="text-muted small"
                      >
                        Sin acciones disponibles
                      </span>
                    </td>
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
      :host {
        display: block;
        margin: 0;
        padding: 0;
        width: 100%;
        min-height: 100vh;
        background: #f8f9fa;
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

      /* ===== DERECHA ===== */

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

      /* ===== LOGOUT ===== */

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

      /* ===== CONTENIDO ===== */

      .main-content {
        min-height: calc(100vh - 56px);
      }

      /* ===== TABLA ===== */

      table {
        width: 100%;
      }

      th {
        white-space: nowrap;
        text-align: center;
        background-color: #8a97ad;
      }

      td {
        vertical-align: middle;
        text-align: center;
        background-color: #0808084b;
      }

      /* ===== BOTONES DE ESTADO ===== */

      td .btn {
        white-space: nowrap;
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
}
