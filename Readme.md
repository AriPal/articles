# Azure AD authentication in Microsoft Teams

Microsoft Teams is a communication platform that serves thousands of users world-wide with the aim to simplify collaboration and management of internal and external processes. Most organizations use Teams because of the chat features it provide, but that is just a small part of what it truly offers.

Did you know that you can build a web application, integrate it with Microsoft Teams, and then share it within your organization to automate time-consuming tasks? Additionally, you integrate it with Microsoft Graph, a Restful web API to interact with Microsofts cloud services such as Excel, SharePoint, OneDrive and so on.

Organizations today have many challenges that can be improved and automated.
For instance one could generate timesheets to Excel from a third-party API to simplify work for the financial team, automate employemeent processes for HR team, sending out weekly news, reports and reminders to employees and so forth. These are few among many good reasons to use Microsoft Graph API endpoint.

![Microsoft Graph](https://docs.microsoft.com/en-us/graph/images/microsoft-graph-dataconnect-connectors-800.png)

As described in by Microsoft - Microsoft Graph is the gateway to data and intellegence in Microsoft 365 (Outloook, SharePoint, OneDrive, Office). A single endpoint `https://graph.microsoft.com` to access data for building and automating apps for organizations and consumers that interact with millions of users. If you are interested and want to try some of the MS Graph endpoints, check out [Microsoft Grap Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer/preview) (UI based platform for MS Graph endpoints).

## Goal
In this article, we'll create a simple web application that generates an Excel file on users OneDrive account. For this to happen we must:
1. Authenticate user in Azure Active Directory (AAD)
2. Get `access_token` from authorization endpoint
3. Then use `access_token` to further access Microsoft Graph endpoint

> What is Azure Actice Directory (AAD), and why do we need it? In order to secure our data from a unauthorized users, we must ensure two things; they belong to the correct domain (company), and consent to provided access permisions like (User.Read, File.Read).

This article won't cover how to integrate a web application with Microsoft Teams using App Studio. Here's a nice article written by Pär that covers it very well.

Packages you need (latest version)
* Microsoft Teams SDK
* Active Directory Authentication Library (ADAL) SDK

The aim of the article is to generate Excel files, and store them on users OneDrive. In order to generate an Excel file, we need to access Microsoft Graph, and in order to access it we must first authenticate the user against Azure AD. So there are few steps before we can end


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

In order to access functions like `login`, `logout`, `getCachedUser`, `getCachedToken` and so forth, we need to create an `AutenticationContext` and pass `config` as argument. You can see this as our first step towards authentication.

```ts
let authContext = new AuthenticationContext(config);
```


## Step 5 - Check if user is logged in

One would think that since the content page is shown in Microsoft Teams, and user has already logged-in in Teams, the user is cached. Unfurtenately that is not the case. As mentioned earlier, our application is an independet app that requires authentication. So if we put our code inside `microsoftTeams.getContext({here...})`, we make sure that the code only runs in Teams.

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






