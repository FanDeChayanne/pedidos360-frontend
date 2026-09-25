import {
  Component,
  OnInit,
  ChangeDetectorRef
} from '@angular/core';

import { CommonModule } from '@angular/common';
import { RouterLink } from '@angular/router';
import { AuthService } from '../core/auth.service';

@Component({
  selector: 'app-orders',
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <div style="padding:30px; font-family:Arial">

      <h2>Pedidos</h2>

      <a routerLink="/dashboard">
        ← Volver al Dashboard
      </a>

      <div
        *ngIf="esCliente"
        style="margin-top:20px"
      >
        <h3>Mis pedidos (Cliente)</h3>

        <p>
          Aquí el cliente ve solo sus pedidos
          y puede crear uno nuevo.
        </p>

        <button>
          Crear pedido (mock)
        </button>
      </div>

      <div
        *ngIf="esOperador || esAdmin"
        style="margin-top:20px"
      >
        <h3>
          Gestión de pedidos (Operador / Admin)
        </h3>

        <p>
          Aquí se ven todos los pedidos
          y se pueden cambiar estados:
        </p>

        <ul>
          <li>
            CREADO → ACEPTADO → EN_PREPARACIÓN
            → DESPACHADO → ENTREGADO / CANCELADO
          </li>
        </ul>

        <p>
          <strong>Regla:</strong>
          No se puede DESPACHAR sin ACEPTAR primero.
        </p>
      </div>

    </div>
  `
})
export class OrdersComponent implements OnInit {

  esAdmin = false;
  esOperador = false;
  esCliente = false;

  constructor(
    private auth: AuthService,
    private cdr: ChangeDetectorRef
  ) {}

  async ngOnInit() {

    const roles =
      await this.auth.getRoles();

    this.esAdmin =
      roles.includes('ROLE_ADMINISTRADOR');

    this.esOperador =
      roles.includes('ROLE_OPERADOR');

    this.esCliente =
      roles.includes('ROLE_CLIENTE');

    console.log(
      'Roles en Pedidos:',
      roles
    );

    console.log(
      'Acceso Admin:',
      this.esAdmin
    );

    console.log(
      'Acceso Operador:',
      this.esOperador
    );

    console.log(
      'Acceso Cliente:',
      this.esCliente
    );

    this.cdr.detectChanges();
  }
}

