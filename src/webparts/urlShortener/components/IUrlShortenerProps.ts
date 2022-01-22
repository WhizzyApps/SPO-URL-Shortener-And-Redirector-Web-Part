import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUrlShortenerProps {
    context: WebPartContext;
    idLength: number;
    lookupList: {id: {}, title: String};
}
