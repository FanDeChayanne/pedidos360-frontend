export const environment = {
  production: true,
  msalConfig: {
    auth: {
      clientId: '25380059-284a-408a-aa5d-22efa594c5dc',
      authority: 'https://login.microsoftonline.com/bdda1d7c-134f-4a96-a337-c592e586a963',
      // Se calcula solo: CloudFront / Amplify / localhost
      redirectUri: window.location.origin,
      postLogoutRedirectUri: window.location.origin
    }
  },
  apiConfig: {
    scopes: ['api://4ea351c0-ef77-41aa-acac-dd63f5d0648f/access_as_user'],
    // El frontend NUNCA llama al BFF (54.146.210.63). Siempre API Gateway.
    uri: 'https://l3rovahkmg.execute-api.us-east-1.amazonaws.com',
    apiAppId: '4ea351c0-ef77-41aa-acac-dd63f5d0648f'
  }
};