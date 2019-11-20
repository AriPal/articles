import React from 'react';
import Home from './Home';
import { BrowserRouter as Router, Route, Link } from 'react-router-dom';
import Authorization from './Authorization';

const Navigation: = () => {
    return (
        <Router>
            <div>
                <section className="hero is-dark">
                    <div className="hero-body">
                        <div className="container">
                            <h1 className="title">
                                Authentication and Authorization Flow in Tabs
                            </h1>
                            <h2 className="subtitle is-6">
                                A simple covering the process of authentication and authorization in order to access Microsoft Graph in MS Teams.
                            </h2>
                        </div>
                    </div>
                </section>

                <Route path="/" exact component={Home}></Route>
                <Route path="/tab/authorization" exact component={Authorization}></Route>
            </div>
        </Router>
    );
}

export default Navigation;
