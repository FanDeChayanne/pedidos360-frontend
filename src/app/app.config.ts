import { ApplicationConfig, APP_INITIALIZER } from "@angular/core";
import { provideRouter } from "@angular/router";
import {
  provideHttpClient,
  withInterceptorsFromDi,
  HTTP_INTERCEPTORS,
} from "@angular/common/http";
import {
  MSAL_INSTANCE,
  MSAL_GUARD_CONFIG,
  MSAL_INTERCEPTOR_CONFIG,
  MsalGuard,
  MsalInterceptor,
  MsalBroadcastService,
  MsalService,
  MsalGuardConfiguration,
  MsalInterceptorConfiguration,
} from "@azure/msal-angular";
import {
  IPublicClientApplication,
  InteractionType,
  PublicClientApplication,
} from "@azure/msal-browser";
import { routes } from "./app.routes";
import { msalConfig, loginRequest, protectedResources } from "./auth-config";

export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication(msalConfig);
}

export function initializeMSAL(msalInstance: IPublicClientApplication) {
  return () => msalInstance.initialize();
}

export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,

    authRequest: loginRequest,

    loginFailedRoute: "/login-failed",
  };
}

export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();

  // se define que la API requiere un token
  protectedResourceMap.set(
    `${protectedResources.apiPedidos360.endpoint}/*`,
    protectedResources.apiPedidos360.scopes,
  );

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap,
  };
}

export const appConfig: ApplicationConfig = {
  providers: [
    provideRouter(routes),

    provideHttpClient(withInterceptorsFromDi()),

    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true,
    },

    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory,
    },

    {
      provide: APP_INITIALIZER,
      useFactory: initializeMSAL,
      deps: [MSAL_INSTANCE],
      multi: true,
    },

    {
      provide: MSAL_GUARD_CONFIG,
      useFactory: MSALGuardConfigFactory,
    },

    {
      provide: MSAL_INTERCEPTOR_CONFIG,
      useFactory: MSALInterceptorConfigFactory,
    },

    MsalService,

    MsalGuard,

    MsalBroadcastService,
  ],
};
