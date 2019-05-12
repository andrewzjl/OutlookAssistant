import 'office-ui-fabric-react/dist/css/fabric.min.css';
import App from './components/App';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

initializeIcons();

let isOfficeInitialized = false;

const title = 'Contoso Task Pane Add-in';

const render = (Component) => {
    ReactDOM.render(
        <AppContainer>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </AppContainer>,
        document.getElementById('container')
    );
};

/* Render application after Office initializes */
Office.initialize = () => {
    isOfficeInitialized = true;
    render(App);

    const { mailbox, displayLanguage } = Office.context || {}
    const { item } = mailbox || {}
  
    if (mailbox.addHandlerAsync) {
      mailbox.addHandlerAsync(Office.EventType.ItemChanged, selectedEmailItemDidChange)
    }
    getBody(item).then((body)=>{
        console.info(body)
        //window.location = window.location
    })
};


const selectedEmailItemDidChange = () => {
    const { item } = Office.context.mailbox || {}
    getBody(item).then((body)=>{
        console.info(body)
        //window.location = window.location
    })
  }
  

async function getBody (item)  {
    try {
        return await new Promise((resolve, reject) => {
            item.body.getAsync('text', (result) => {
              if (result.status === 'succeeded') {
                console.log(result.value)
                return resolve(result.value); // updated as suggested by Mavi Domates
              } else {
                console.error(result.error)
                return reject(result.error);
              }
            })
          })
    } catch (error) {
        alert(error)
    }
}

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
    module.hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}