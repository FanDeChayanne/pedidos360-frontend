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

    <nav class="top-navbar">

      <div class="nav-container">

        <div class="nav-left">

          <a
            class="nav-item"
            routerLink="/dashboard"
          >
            <i class="bi bi-house-door-fill"></i>
            Dashboard
          </a>

          <a
            class="nav-item"
            routerLink="/orders"
          >
            <i class="bi bi-cart3"></i>
            Pedidos
          </a>

          <a
            class="nav-item active"
            *ngIf="esAdmin || esOperador"
            routerLink="/catalog"
          >
            <i class="bi bi-journal-text"></i>
            Catálogo
          </a>

          <!-- Usuarios (solo Admin) -->

          <a
            class="nav-item"
            *ngIf="esAdmin"
            routerLink="/usuarios"
          >
            <i class="bi bi-people"></i>
            Usuarios
          </a>

          <a
            class="nav-item"
            href="javascript:void(0)"
          >
            <i class="bi bi-box"></i>
            TEST
          </a>

        </div>

        <!-- ========== USUARIO ========== -->

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

          <button
            class="btn-logout"
            (click)="logout()"
          >
            <i class="bi bi-box-arrow-right"></i>
            Cerrar sesión
          </button>

        </div>

      </div>

    </nav>


    <!-- ========== CONTENIDO ========== -->

    <main class="main-content">

      <div class="p-4">

        <h2 class="mb-3">
          Catálogo de Productos
        </h2>


        <!-- ========== CREAR PRODUCTO - SOLO ADMIN ========== -->

        <div
          *ngIf="esAdmin"
          class="card border-0 shadow-sm mb-4"
        >

          <div class="card-body">

            <h5 class="card-title">
              <i class="bi bi-plus-circle me-2"></i>
              Crear producto
            </h5>

            <p class="card-text text-muted">
              Agrega un nuevo producto al catálogo.
            </p>


            <div class="row g-3">

              <!-- NOMBRE -->

              <div class="col-md-4">

                <label class="form-label">
                  Nombre del producto
                </label>

                <input
                  type="text"
                  class="form-control"
                  [(ngModel)]="nuevoProducto.nombre"
                  name="nombreProducto"
                  placeholder="Ej: Producto nuevo"
                >

              </div>


              <!-- PRECIO -->

              <div class="col-md-4">

                <label class="form-label">
                  Precio
                </label>

                <input
                  type="number"
                  class="form-control"
                  min="0"
                  [(ngModel)]="nuevoProducto.precio"
                  name="precioProducto"
                  placeholder="Ej: 5000"
                >

              </div>


              <!-- STOCK INICIAL -->

              <div class="col-md-4">

                <label class="form-label">
                  Stock inicial
                </label>

                <input
                  type="number"
                  class="form-control"
                  min="0"
                  [(ngModel)]="nuevoProducto.stock"
                  name="stockProducto"
                  placeholder="Ej: 10"
                >

              </div>

            </div>


            <!-- BOTÓN -->

            <div class="mt-3">

              <button
                type="button"
                class="btn btn-primary"
                (click)="crearProducto()"
              >

                <i class="bi bi-plus-circle me-2"></i>

                Crear producto

              </button>

            </div>

          </div>

        </div>


        <!-- ========== GESTIÓN DE STOCK ========== -->

        <div
          *ngIf="esAdmin || esOperador"
          class="card border-0 shadow-sm mb-4"
        >

          <div class="card-body">

            <div
              class="d-flex justify-content-between align-items-center mb-3"
            >

              <div>

                <h5 class="card-title mb-1">
                  Gestión de stock
                </h5>

                <p class="text-muted mb-0">
                  Agrega unidades al stock de los productos.
                </p>

              </div>


              <span class="badge bg-secondary">
                {{ productos.length }} productos
              </span>

            </div>


            <!-- SIN PRODUCTOS -->

            <div
              *ngIf="productos.length === 0"
              class="alert alert-secondary mb-0"
            >
              No hay productos registrados.
            </div>


            <!-- TABLA DE PRODUCTOS -->

            <div
              *ngIf="productos.length > 0"
              class="table-responsive"
            >

              <table
                class="table table-hover align-middle mb-0"
              >

                <thead>

                  <tr>

                    <th>
                      ID
                    </th>

                    <th>
                      Producto
                    </th>

                    <th>
                      Precio
                    </th>

                    <th>
                      Stock actual
                    </th>

                    <th>
                      Cantidad a agregar
                    </th>

                    <th>
                      Acción
                    </th>

                  </tr>

                </thead>


                <tbody>

                  <tr
                    *ngFor="let producto of productos"
                  >

                    <td>
                      {{ producto.id }}
                    </td>


                    <td>
                      <strong>
                        {{ producto.nombre }}
                      </strong>
                    </td>


                    <td>
                      $ {{ producto.precio }}
                    </td>


                    <td>

                      <span
                        class="badge"
                        [ngClass]="{
                          'bg-danger': producto.stock === 0,
                          'bg-warning text-dark': producto.stock > 0 && producto.stock <= 10,
                          'bg-success': producto.stock > 10
                        }"
                      >
                        {{ producto.stock }}
                      </span>

                    </td>


                    <td style="max-width: 180px;">

                      <input
                        type="number"
                        class="form-control"
                        min="1"
                        [(ngModel)]="cantidadesStock[producto.id]"
                        [name]="'stock-' + producto.id"
                      />

                    </td>


                    <td>

                      <button
                        type="button"
                        class="btn btn-success btn-sm"
                        (click)="aumentarStock(producto)"
                      >

                        <i class="bi bi-plus-circle"></i>

                        Agregar stock

                      </button>

                    </td>

                  </tr>

                </tbody>

              </table>

            </div>

          </div>

        </div>


        <!-- ========== SIN PERMISOS ========== -->

        <div
          *ngIf="!esAdmin && !esOperador"
          class="alert alert-danger"
        >
          No tienes permiso para administrar el catálogo.
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


      /* ========== NAVBAR ========== */

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


      /* ========== USUARIO ========== */

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


      /* ========== LOGOUT ========== */

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


      /* ========== CONTENIDO ========== */

      .main-content {
        min-height: calc(100vh - 56px);
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

}