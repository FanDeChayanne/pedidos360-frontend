export const environment = {
  production: false,
  msalConfig: {
    auth: {
      // SPA que YA te funciona en local. No cambies esto salvo que Entra lo pida.
      clientId: '25380059-284a-408a-aa5d-22efa594c5dc',
      authority: 'https://login.microsoftonline.com/bdda1d7c-134f-4a96-a337-c592e586a963',
      redirectUri: window.location.origin,
      postLogoutRedirectUri: window.location.origin
    }
  },
  apiConfig: {
    // Access token para el BFF (audience del JWT Authorizer de API Gateway)
    scopes: ['api://4ea351c0-ef77-41aa-acac-dd63f5d0648f/access_as_user'],
    uri: 'http://localhost:8080',
    apiAppId: '4ea351c0-ef77-41aa-acac-dd63f5d0648f'
  }
};