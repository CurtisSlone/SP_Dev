import * as React from 'react';
import styles from './RecentProducts.module.scss';
import { IDocumentProps } from './IDocumentProps';
import { DocumentCard,
    DocumentCardActivity,
    DocumentCardPreview,
    DocumentCardTitle,
    DocumentCardType,
} from 'office-ui-fabric-react/lib/DocumentCard';
import {
    CompoundButton,
    IButtonProps
  } from 'office-ui-fabric-react/lib/Button';
  import { ImageFit } from 'office-ui-fabric-react/lib/Image';

export default class DocumentComponent extends React.Component<IDocumentProps,{}> {
    constructor(props: IDocumentProps){
        super(props);
    }

    public render(): React.ReactElement<IDocumentProps>{
        
        return (
            <div className={styles.column}>
                <DocumentCard type={ DocumentCardType.normal } onClickHref='http://google.com' >
                    <img className={styles.imgfit} src={this.props.documentImgURL} />
                    <DocumentCardTitle
                        title={this.props.documentName}
                        shouldTruncate={ true}
                    />
                    <DocumentCardTitle
                        title={this.props.documentDescription}
                        shouldTruncate={ true}
                    />
                </DocumentCard>
            </div>
        );
    }
}