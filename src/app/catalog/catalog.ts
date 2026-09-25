import { Component, OnInit, ChangeDetectorRef } from "@angular/core";

import { CommonModule } from "@angular/common";
import { RouterLink } from "@angular/router";
import { AuthService } from "../core/auth.service";

@Component({
  selector: "app-catalog",
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <div style="padding:30px; font-family:Arial">
      <h2>Catálogo de Productos</h2>

      <a routerLink="/dashboard"> ← Volver al Dashboard </a>

      <div *ngIf="esAdmin || esOperador" style="margin-top:20px">
        <h3>Gestión de catálogo</h3>

        <p>CRUD de productos, precios y stock.</p>

        <p>El stock disminuye al aceptar un pedido.</p>

        <button>Crear producto (mock)</button>

        <button style="margin-left:10px">Editar stock (mock)</button>
      </div>

      <div
        *ngIf="!esAdmin && !esOperador"
        style="
          margin-top:20px;
          color:red
        "
      >
        No tienes permiso para administrar el catálogo.
      </div>
    </div>
  `,
})
export class CatalogComponent implements OnInit {
  esAdmin = false;
  esOperador = false;

  constructor(
    private auth: AuthService,
    private cdr: ChangeDetectorRef,
  ) {}

  async ngOnInit() {
    const roles = await this.auth.getRoles();

    this.esAdmin = roles.includes("ROLE_ADMINISTRADOR");

    this.esOperador = roles.includes("ROLE_OPERADOR");

    console.log("Roles en Catálogo:", roles);

    console.log("Acceso Admin:", this.esAdmin);

    console.log("Acceso Operador:", this.esOperador);

    this.cdr.detectChanges();
  }
}
