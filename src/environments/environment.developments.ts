export const environment = {
  production: false,
  msalConfig: {
    auth: {
      clientId: '42f88b93-ddd9-44c1-9b38-3b35ada854e2',
      authority: 'https://login.microsoftonline.com/bdda1d7c-134f-4a96-a337-c592e586a963',
      redirectUri: 'http://localhost:4200',
      postLogoutRedirectUri: 'http://localhost:4200'
    }
  },
  apiConfig: {
    scopes: ['api://42f88b93-ddd9-44c1-9b38-3b35ada854e2/access_as_user'],
    uri: 'https://URL_DEL_API_GATEWAY_AQUI' // cambiarlo despues
  }
};