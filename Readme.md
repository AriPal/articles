# Azure AD authentication in Microsoft Teams

Microsoft Teams is a communication platform that serves thousands of users world-wide with the aim to simplify collaboration and management of internal and external processes. Most organizations use Microsoft Teams because of the chat features it provide, but that's just a small part of what it truly offers.

Organizations today use many services as Microsoft provide, but know very little of what flexibility and

![Netcompany logo](https://miro.medium.com/max/137/1*CWxqiYnxrOzHP-t8wQWfjQ.png)

_I work as a Software Consultant at Netcompany Norway - an international company with more than 2000 employees in 6 countries that provides valuable expertise and experience to help organizations with digital transformation. Currently, they are exploring new fields and ways to automate time-consuming processes using Microsoft cloude services. One of them is to integrate a web app with Microsoft Teams to reduce time-consuming tasks._


![Microsoft Graph](https://docs.microsoft.com/en-us/graph/images/microsoft-graph-dataconnect-connectors-800.png)

As described by Microsoft - Microsoft Graph is the gateway to data and intellegence in Microsoft 365 (Outloook, SharePoint, Office). A single endpoint `https://graph.microsoft.com` to access data to builde and automate apps for organizations and consumers that interact with millions of users. If you are interested and want to try some of the MS Graph endpoints, check out [Microsoft Grap Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer/preview) (UI based platform for MS Graph endpoints).

## Outcome

The outcome of this article is to show how to store an Excel file on user's OneDrive account following these main steps:

1. Authenticate user in Azure Active Directory (AAD)
2. Send request to autherization endpoint to get `access_token`
3. Use `access_token` to interact with Microsoft Graph

> What is Azure Actice Directory (AAD), and why do we need it? In order to secure our data from a unauthorized users, we must ensure two things; they belong to the correct domain (company), and consent to provided access permisions like (User.Read, File.Read).

This article won't cover how to integrate a web application with Microsoft Teams using App Studio. Here's a nice article written by Pär that covers it very well.

## Prerequisites

* Office 365 Developer subscription
* A deployed web application
* Microsoft Teams

## Packages you need (latest version)

* Microsoft Teams SDK
* Active Directory Authentication Library (ADAL) SDK

The aim of the article is to generate an Excel file when user clicks on a button, and store it on user's OneDrive. In order to generate an Excel file, we need to access Microsoft Graph, and in order to access it we must first authenticate the user against Azure AD. So there are few steps before we can end


## Step 1 - Import NPM packages

The language we are using for this example is TypeScript. If you also want to use TypeScript, you need to install the `@types` packages.

```ts
import AuthenticationContext, { Options } from 'adal-angular';
import * as microsoftTeams from '@microsoft/teams-js';
```

## Step 2 - Initialize Microsoft Teams
The web application (content page) shown in Microsoft Teams is done through `Iframe`, which means it is not fully integrated. So anyone with the url link can easily access our application, and perhaps do some harm. In order to secure our application, we need a way to tell it that it can only authenticate and request data from MS Graph if its inside the Microsoft Teams context, we'll discuss this later.

```ts
microsoftTeams.initialize();
```

## Step 3 - Setup Configuration

```ts
let config: Options = {
    tenant: 'tenant_id', // Found in app registration in azure.portal.com
    clientId: 'client_id', // Found in app registration in azure.portal.com
    redirectUri: 'URI' + '/auth/silent-end', // Important: URL must be registered in Redirect URL otherwise it won't work.
    cacheLocation: "localStorage",
    popUp: true, // A popup form shown only if auth fails
    navigateToLoginRequestUrl: false,
};
```

## Step 4 - Create a configuration context

In order to access functions like `login`, `logout`, `getCachedUser`, `getCachedToken` and so forth, we need to create an `AutenticationContext` and pass `config` as argument. You can see this as our first step towards interacting with authentication.

```ts
let authContext = new AuthenticationContext(config);
```


## Step 5 - Check if user is logged in

Once we have created an `AuthenticationContext(config)`, the next step is check if user is cached. Keep in mind that the whole authentication process is done through `MicrosoftTeams.getContext({...auth process...})`



```ts
microsoftTeams.getContext((context) => {
    let user = authContext.getCachedUser();
    if (user) {
        // Code that runs here means the user has been authenticated
        // Get access_token for MS Graph
    } else {
        // Show login popup
        authContext.login();
    }
}

```

## Step 6 - Handle cache

If the current user is not the same as the one in Teams context, we clear the cache. Doing so, the new user has to login again. The code shown in step 5 will be triggered.

```ts
// Clear cache if expected user is different
if (user.userName !== context.upn) {
    authContext.clearCache();
}
````






