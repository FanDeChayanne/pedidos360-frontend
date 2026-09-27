export const environment = {
  production: false,
  msalConfig: {
    auth: {
      clientId: '25380059-284a-408a-aa5d-22efa594c5dc',
      authority: 'https://login.microsoftonline.com/bdda1d7c-134f-4a96-a337-c592e586a963',
      redirectUri: 'http://localhost:4200',
      postLogoutRedirectUri: 'http://localhost:4200'
    }
  },
  apiConfig: {
    scopes: ['api://4ea351c0-ef77-41aa-acac-dd63f5d0648f/access_as_user'],
    uri: 'http://localhost:8080',
    apiAppId: '4ea351c0-ef77-41aa-acac-dd63f5d0648f'
  }
};