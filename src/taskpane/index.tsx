import 'office-ui-fabric-react/dist/css/fabric.min.css';
import App from './components/App';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Provider } from "react-redux";
import configureStore from "../store";

initializeIcons();

let isOfficeInitialized = false;

const store = configureStore();

const render = (Component) => {
    ReactDOM.render(
        <Provider store={store}>
            <AppContainer>
                <Component isOfficeInitialized={isOfficeInitialized} />
            </AppContainer>
        </Provider>,
        document.getElementById('container')
    );
};

Office.onReady(function(info) {
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
    isOfficeInitialized = true;
    render(App);
});

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
    (module as any).hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}