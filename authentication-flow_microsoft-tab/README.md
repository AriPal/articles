## Note

This code example is mainly for showing how Authentication and Authorization work for a content page in Microsoft Teams.

If you want to test this application on Microsoft Teams, you need to create an application in App Studio, and then link the content page to an existing domain.

Here's a quick [guide](https://medium.com/@paumadregis/custom-microsoft-teams-applications-the-easy-way-6da0a5975336) showing how to public a web app on Teams by Pär Joona.

> If you are using Netlify to host your web app, remember to add `/*` wildcard in _redirects file, otherwise when we try to authenticate or authorize user it will return 'page not found`.