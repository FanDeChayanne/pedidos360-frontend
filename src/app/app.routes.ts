import { Routes } from '@angular/router';

import { MsalGuard } from '@azure/msal-angular';

import { LoginComponent } from './login/login';
import { DashboardComponent } from './dashboard/dashboard';
import { OrdersComponent } from './orders/orders';
import { CatalogComponent } from './catalog/catalog';
import { LoginFailedComponent } from './login-failed/login-failed';
import { ProfileComponent } from './profile/profile';

import { roleGuard } from './core/role.guard';

export const routes: Routes = [

  // LOGIN
  {
    path: 'login',
    component: LoginComponent
  },

  
  // DASHBOARD
  {
    path: 'dashboard',
    component: DashboardComponent,
    canActivate: [
      MsalGuard
    ]
  },

  // PEDIDOS
  {
    path: 'orders',
    component: OrdersComponent,
    canActivate: [
      MsalGuard
    ]
  },

  // CATÁLOGO
  {
    path: 'catalog',
    component: CatalogComponent,
    canActivate: [
      MsalGuard,
      roleGuard([
        'ROLE_ADMINISTRADOR',
        'ROLE_OPERADOR'
      ])
    ]
  },

  
  // PERFIL
  {
    path: 'profile',
    component: ProfileComponent,
    canActivate: [
      MsalGuard
    ]
  },

  
  // LOGIN FALLIDO
  {
    path: 'login-failed',
    component: LoginFailedComponent
  },

  
  // RUTA PRINCIPAL
  {
    path: '',
    redirectTo: 'login',
    pathMatch: 'full'
  },

  
  // RUTA DESCONOCIDA
  {
    path: '**',
    redirectTo: 'login'
  }
];

