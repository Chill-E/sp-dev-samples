import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Birthday } from './components/Birthday/Birthday';

const webparts: NodeListOf<Element> = document.getElementsByClassName('webpart-birthday');
for (let i: number = 0; i < webparts.length; i++) {
    // Get the data property from the Element
    const description: string = webparts[i].getAttribute('data-description').toString();
    const siteUrl: string = webparts[i].getAttribute('data-site-url').toString();
    const listName: string = webparts[i].getAttribute('data-list-name').toString();

    ReactDOM.render(
        <Birthday description={description} siteUrl={siteUrl} listName={listName} />,
        webparts[i]
    );
}