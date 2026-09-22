import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterLink } from '@angular/router';

@Component({
  selector: 'app-login-failed',
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <div style="padding: 40px; text-align: center;">
      <h2>Error al iniciar sesion</h2>
      <p>Hubo un problema con la autenticacion.</p>
      <a routerLink="/">Volver al inicio</a>
    </div>
  `
})
export class LoginFailedComponent {}