import AuthenticationContext, { Options } from 'adal-angular';
import * as microsoftTeams from '@microsoft/teams-js';
import * as XLSX from 'xlsx';

const Home: React.FC = () => {

    microsoftTeams.initialize();

    let config: Options = {
        tenant: '{ tenant_id }',
        clientId: '{ client_id }',
        redirectUri: window.location.origin + '/auth/silent-end',
        cacheLocation: "localStorage",
        popUp: true,
        navigateToLoginRequestUrl: false,
    };

    let authContext = new AuthenticationContext(config);

    microsoftTeams.getContext((context) => {

        let user = authContext.getCachedUser();
        if (user == null) {
            authContext.login();
        }

        // Clear cache if expected user is not the same as cached user
        if (user.userName !== context.upn) {
            authContext.clearCache();
        }

    })

    function generateExcelFile() {
        microsoftTeams.getContext((context) => {
            authContext.acquireToken(config.clientId, function (errorDesc, token, error) {
                if (error) {
                    // If Show sign in button
                }
                else {
                    let queryParams = {
                        client_id: config.clientId,
                        response_type: "token",
                        response_mode: "fragment",
                        scope: "https://graph.microsoft.com/User.Read openid",
                        redirect_uri: window.location.origin + '/auth/autherization',
                        nonce: 1234,
                        state: 5687,
                        login_hint: context.loginHint,
                    };

                    let authorizeEndpoint = `https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/authorize?${toQueryString(queryParams)}`;
                    window.location.assign(authorizeEndpoint);

                    // Get cached access_token
                    let access_token = window.localStorage.getItem('access_token');

                    if (access_token) {

                        // Basic fetch API request
                        const headers = new Headers();
                        headers.append('Authorization', `Bearer ${access_token}`);
                        headers.append('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

                        var graphEndpoint = "https://graph.microsoft.com/v1.0/drive/root:/mellomprosjekt/Yabadabadoo.xlsx:/content";

                        // Create Excel File
                        const book = XLSX.utils.book_new();
                        const data = [];
                        data.push(['label1', 'label2']);
                        data.push(['data1', 'data2']);
                        const sheet = XLSX.utils.aoa_to_sheet(data);
                        XLSX.utils.book_append_sheet(book, sheet, 'sheet1');
                        let workBook = XLSX.write(book, { bookType: 'xlsx', type: 'array' });

                        console.log('workBook', workBook);
                        let init: RequestInit = {
                            method: 'PUT',
                            headers: headers,
                            body: workBook,
                        };

                        // Run request
                        fetch(graphEndpoint, init)
                            .then(resp => resp.json())
                            .then(data => console.log(data))

                    }
                }
            });
        })
    }


    return (
        <div style={{ padding: 40 }} >
            <button className="button is-dark" style={{ marginBottom: 20 }} onClick={() => generateExcelFile()}>Generate Excel File</button>
            <p className="subtitle is-5"><strong>User:</strong> {userProfile.displayName}</p>
            <p className="subtitle is-5"><strong>Job title:</strong> {userProfile.jobTitle}</p>
            <p className="subtitle is-5"><strong>Office location:</strong> {userProfile.officeLocation}</p>
            <p className="subtitle is-5"><strong>Email:</strong> {userProfile.userPrincipalName}</p>
        </div>
    );
}


export default Home;
