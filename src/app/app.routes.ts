import { Routes } from '@angular/router';
import { MsalGuard } from '@azure/msal-angular';
import { HomeComponent } from './home/home';
import { ProfileComponent } from './profile/profile';
import { LoginFailedComponent } from './login-failed/login-failed';

export const routes: Routes = [

    {
        path: '',
        component: HomeComponent
    },
    {
        path: 'profile',
        component: ProfileComponent,
        canActivate: [MsalGuard]
    },
    {
        path: 'login-failed',
        component: LoginFailedComponent
    },
    {
        path: '**',
        redirectTo: ''
    }
];