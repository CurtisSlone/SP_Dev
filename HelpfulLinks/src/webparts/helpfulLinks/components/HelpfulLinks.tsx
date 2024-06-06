import * as React from 'react';
import styles from './HelpfulLinks.module.scss';
import { IHelpfulLinksProps } from './IHelpfulLinksProps';
import { DocumentCard,
  DocumentCardTitle,
  DocumentCardType,
} from 'office-ui-fabric-react/lib/DocumentCard';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelpfulLinks extends React.Component<IHelpfulLinksProps, {}> {
  constructor(props: IHelpfulLinksProps){
    super(props);

  }

  
  public render(): React.ReactElement<IHelpfulLinksProps> {
    
    const links: any[] = [];
    for( let i: number = 0; i < this.props.linkCount; i++)
      links.push(
        <div className={styles.column} >
          <DocumentCard 
            type={ DocumentCardType.normal }
            onClickHref={this.props.linkUrls[i]}
            className={styles.docCard}
          >
            <h3 className={styles.largeTitle}>{this.props.linkNames[i]}</h3>
          </DocumentCard>
        </div>
      );
    return (
      <div className={ styles.helpfulLinks }>
        <div className={ styles.container }>
          <div className={ styles.row }>
                {links}
          </div>
        </div>
      </div>
    );
  }
}
