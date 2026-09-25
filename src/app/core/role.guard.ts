import { inject } from '@angular/core';
import { CanActivateFn } from '@angular/router';
import { Router } from '@angular/router';
import { AuthService } from './auth.service';

export function roleGuard(
  allowedRoles: string[]
): CanActivateFn {

  return async () => {

    console.log('=== ROLE GUARD EJECUTADO ===');

    const auth = inject(AuthService);
    const router = inject(Router);

    console.log('¿Está logueado?', auth.isLoggedIn());

    if (!auth.isLoggedIn()) {
      console.log('No está logueado → /login');

      return router.createUrlTree([
        '/login'
      ]);
    }

    const roles = await auth.getRoles();

    console.log('Roles detectados:', roles);
    console.log('Roles permitidos:', allowedRoles);

    const hasPermission =
      roles.some(
        role => allowedRoles.includes(role)
      );

    console.log('¿Tiene permiso?', hasPermission);

    if (hasPermission) {
      console.log('ACCESO PERMITIDO');
      return true;
    }

    console.log('ACCESO DENEGADO → /dashboard');

    return router.createUrlTree([
      '/dashboard'
    ]);
  };
}

