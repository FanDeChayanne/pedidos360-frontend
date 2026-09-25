import { Component, OnInit } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { CommonModule } from '@angular/common';
import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import { InteractionStatus } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet],
  template: `<router-outlet></router-outlet>`,
  styles: []
})
export class AppComponent implements OnInit {

  constructor(
    private authService: MsalService,
    private msalBroadcastService: MsalBroadcastService
  ) {}

  ngOnInit(): void {

    this.authService.handleRedirectObservable({
      navigateToLoginRequestUrl: false
    }).subscribe(result => {
      console.log('MSAL redirect procesado:', result);
    });

    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) =>
          status === InteractionStatus.None
        )
      )
      .subscribe(() => {
        console.log(
          'MSAL terminó interacción. Cuentas:',
          this.authService.instance.getAllAccounts()
        );
      });
  }
}