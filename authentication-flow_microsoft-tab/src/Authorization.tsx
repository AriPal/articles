import React from 'react';
import { useHistory } from 'react-router-dom'
import qs from 'querystring';


const Authorization = (props: any) => {
    let history = useHistory();
    let querystring = qs.parse(props.location.hash)
    let access_token = querystring['#access_token'].toString();
    if(access_token) {
        window.localStorage.setItem('access_token', access_token);
        history.push('/'); // Redirect to home page
    }
    return (
        <div>
            <h1>This component is only used to capture (cache) the access_token from Azure AD, and then redirect back to home page</h1>
        </div>
    );
}

export default Authorization;
