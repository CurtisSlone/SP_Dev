import * as React from 'react';
import { IDocumentProps } from './IDocumentProps';

export default class DocumentComponent extends React.Component<IDocumentProps,{}> {
    constructor(props: IDocumentProps){
        super(props);
    }

    public render(): React.ReactElement<IDocumentProps>{

        return (
        
            <li>
                <div>{this.props.documentId}</div>
                <div>{this.props.documentName}</div>
                <div>{this.props.documentClassification}</div>
                <div>{this.props.documentDescription}</div>
                <div>{this.props.documentImgURL}</div>
            </li>
        );
    }
}