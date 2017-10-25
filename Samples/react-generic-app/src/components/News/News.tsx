import * as React from 'react';
import styles from './News.module.scss';
import { css } from 'office-ui-fabric-react';
import { INewsProps } from './INewsProps';
import { INewsState } from './INewsState';

export class News extends React.Component<INewsProps, INewsState> {
  constructor(props?: INewsProps, context?: any) {
    super();
    this.state = {
      listTitles: [],
      loadingLists: false,
      error: null
    };
  }

  componentDidMount() {
    this.getNewsList();
  }

  public render(): React.ReactElement<INewsProps> {
    const titles: JSX.Element[] = this.state.listTitles.map((item: string, key: number, listTitles: string[]): JSX.Element => {
      return <li key={key}><a ref="#">{item["ContentTypeId"]}</a> {item["Title"]}</li>;
    });
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={css('ms-Grid-row ms-bgColor-teal ms-fontColor-white', styles.row)}>
            <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
              <span className='ms-font-xl ms-fontColor-white'>
                Welcome to SharePoint!
              </span>
              <p className='ms-font-l ms-fontColor-white'>
                Building experiences with web stack and modern tooling
              </p>
              <p className='ms-font-l ms-fontColor-white'>
                {this.props.description}
              </p>
              <div className='test'>
                {this.state.loadingLists &&
                  <span>Loading lists...</span>}
                {this.state.error &&
                  <span>An error has occurred while loading lists: {this.state.error}</span>}
                {this.state.error === null && titles &&
                  <ul>
                    {titles}
                  </ul>}
              </div>
              <a className={css('ms-Button', styles.button)} href='https://dev.office.com/sharepoint'>
                <span className='ms-Button-label'>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private getNewsList() {
    var reactHandler = this;
    reactHandler.setState({
      loadingLists: true,
      listTitles: [],
      error: null
    });

    var spRequest = new XMLHttpRequest();
    spRequest.open('GET', "http://sp13dev:81/sites/zerhusen/_api/web/lists/getbytitle('news')/items", true);
    spRequest.setRequestHeader("Accept", "application/json;odata=verbose");

    spRequest.onreadystatechange = function () {

      if (spRequest.readyState === 4 && spRequest.status === 200) {
        var result = JSON.parse(spRequest.responseText);
        // var resultTitles = [];
        // forearch(var item in result["d"]["result"]){
        //   resultTitles.push(item);
        // })
        reactHandler.setState({
          listTitles: result.d.results,
          loadingLists: false
        });
      }
      else if (spRequest.readyState === 4 && spRequest.status !== 200) {
        console.log('Error Occured !');
        reactHandler.setState({
          error: "Error occured"
        })
      }
    };
    spRequest.send();
  }
}