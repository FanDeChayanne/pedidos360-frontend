import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MsalService } from '@azure/msal-angular';
import { RouterLink } from '@angular/router';

@Component({
  selector: 'app-profile',
  standalone: true,
  imports: [CommonModule, RouterLink],
  templateUrl: './profile.html',
  styleUrl: './profile.scss'
})
export class ProfileComponent implements OnInit {
  profile: any = null;
  roles: string[] = [];
  claims: any = {};

  constructor(private authService: MsalService) {}

  ngOnInit(): void {
    const account = this.authService.instance.getActiveAccount();
    
    if (!account) {
      // Si no hay cuenta activa, toma la primera
      const accounts = this.authService.instance.getAllAccounts();
      if (accounts.length > 0) {
        this.authService.instance.setActiveAccount(accounts[0]);
      }
    }

    const activeAccount = this.authService.instance.getActiveAccount();
    if (activeAccount) {
      this.profile = activeAccount;
      const idClaims = activeAccount.idTokenClaims as any;
      this.claims = idClaims || {};
      this.roles = idClaims?.roles || idClaims?.role || [];
    }
  }
}