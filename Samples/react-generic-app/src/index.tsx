import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { News } from './components/News/News';

const webparts: NodeListOf<Element> = document.getElementsByClassName('webpart-script-example');
for (let i: number = 0; i < webparts.length; i++) {
    // Get the data property from the Element
    const description: string = webparts[i].getAttribute('data-description').toString();

    ReactDOM.render(
        <News description={description} siteUrl={"http://sp13dev:81/sites/zerhusen/SitePages/react-app.aspx"}/>,
        webparts[i]
    );
}