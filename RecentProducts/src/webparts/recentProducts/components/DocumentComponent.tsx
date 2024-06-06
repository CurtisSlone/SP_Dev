import * as React from 'react';
import styles from './RecentProducts.module.scss';
import { IDocumentProps } from './IDocumentProps';
import { DocumentCard,
    DocumentCardTitle,
    DocumentCardType,
} from 'office-ui-fabric-react/lib/DocumentCard';


export default class DocumentComponent extends React.Component<IDocumentProps,{}> {
    constructor(props: IDocumentProps){
        super(props);
    }

    public render(): React.ReactElement<IDocumentProps>{
        
        return (
                <DocumentCard type={ DocumentCardType.normal } >
                    <img className={styles.imgfit} src={this.props.documentImgURL} />
                    <DocumentCardTitle
                        title={this.props.documentName}
                        shouldTruncate={ false }
                    />
                    <DocumentCardTitle
                        title={this.props.documentDescription}
                        shouldTruncate={ true }
                    />
                </DocumentCard>
        );
    }
}