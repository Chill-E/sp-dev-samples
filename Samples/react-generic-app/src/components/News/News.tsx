import * as React from "react";
import styles from "./News.module.scss";
import { css } from "office-ui-fabric-react";
import { INewsProps } from "./INewsProps";
import { INewsState } from "./INewsState";
require("sp-init");
require("microsoft-ajax");
require("sp-runtime");
require("sharepoint");

export class News extends React.Component<INewsProps, INewsState> {
  constructor(props?: INewsProps, context?: any) {
    super();
    this.state = {
      listTitles: [],
      newsList: [],
      loadingLists: false,
      error: null
    };

    this.getMoreNews = this.getMoreNews.bind(this);
  }

  componentDidMount() {
    this.getNewsListRest();
    // this.getNewsListCsom();
  }

  public render(): React.ReactElement<INewsProps> {
    const titles: JSX.Element[] = this.state.listTitles.map(
      (item: string, key: number, listTitles: string[]): JSX.Element => {
        return (
          <div className={styles.newsItem} key={key}>
            <h2 className={styles.head}>{item["Title"]}</h2>
            <p dangerouslySetInnerHTML={{ __html: item["Body"] }} />
          </div>
        );
      }
    );
    // const news: JSX.Element[] = this.state.newsList.map((value: string, key: number, newsList: string[]): JSX.Element => {
    //   return <li key={key}>{value}</li>;
    // });
    return (
      <div className={styles.news}>
        <div className={styles.container}>
          <div
            className={css(
              "ms-Grid-row ms-bgColor-teal ms-fontColor-white",
              styles.row
            )}
          >
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <h1 className={styles.header}>{this.props.description} </h1>
              <div className="news-list-container">
                {this.state.loadingLists && <span>Loading lists...</span>}
                {this.state.error && (
                  <span>
                    An error has occurred while loading news: {this.state.error}
                  </span>
                )}
                {this.state.error === null &&
                  titles && <div className={styles.newsListItem}>{titles}</div>}
              </div>
              <div>
                <ul />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private getNewsListRest() {
    var reactHandler = this;
    reactHandler.setState({
      loadingLists: true,
      listTitles: [],
      error: null
    });

    var spRequest = new XMLHttpRequest();
    var now = new Date().toISOString();
    var listUrl =
      this.props.siteUrl +
      "/_api/web/lists/getbytitle('" +
      this.props.listName +
      "')/items?$filter=(Expires gt datetime'" +
      now +
      "') and (StartDate le datetime'" +
      now +
      "')&$top=5&$orderby=Top desc, StartDate asc";
    spRequest.open("GET", listUrl, true);
    spRequest.setRequestHeader("Accept", "application/json;odata=verbose");

    spRequest.onreadystatechange = function() {
      if (spRequest.readyState === 4 && spRequest.status === 200) {
        var result = JSON.parse(spRequest.responseText);
        reactHandler.setState({
          listTitles: result.d.results,
          loadingLists: false
        });
      } else if (spRequest.readyState === 4 && spRequest.status !== 200) {
        console.log("Error Occured !");
        reactHandler.setState({
          error: "Error occured"
        });
      }
    };
    spRequest.send();
  }

  private getNewsListCsom(): void {
    this.setState({
      loadingLists: true,
      listTitles: [],
      error: null
    });

    const context: SP.ClientContext = new SP.ClientContext(this.props.siteUrl);
    const lists: SP.ListCollection = context.get_web().get_lists();
    context.load(lists, "Include(Title)");
    context.executeQueryAsync(
      (sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
        const listEnumerator: IEnumerator<SP.List> = lists.getEnumerator();

        const newsTitles: string[] = [];
        while (listEnumerator.moveNext()) {
          const list: SP.List = listEnumerator.get_current();
          newsTitles.push(list.get_title());
        }

        this.setState(
          (prevState: INewsState, props: INewsProps): INewsState => {
            prevState.newsList = newsTitles;
            prevState.loadingLists = false;
            return prevState;
          }
        );
      },
      (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
        this.setState({
          loadingLists: false,
          newsList: [],
          error: args.get_message()
        });
      }
    );
  }

  private getMoreNews(): void {
    var reactHandler = this;
    reactHandler.setState({
      loadingLists: true,
      listTitles: [],
      error: null
    });

    var spRequest = new XMLHttpRequest();
    var now = new Date().toISOString();
    var listUrl =
      this.props.siteUrl +
      "/_api/web/lists/getbytitle('" +
      this.props.listName +
      "')/items?$filter=(Expires gt datetime'" +
      now +
      "') and (StartDate le datetime'" +
      now +
      "')&$orderby=Top desc, StartDate asc";
    spRequest.open("GET", listUrl, true);
    spRequest.setRequestHeader("Accept", "application/json;odata=verbose");

    spRequest.onreadystatechange = function() {
      if (spRequest.readyState === 4 && spRequest.status === 200) {
        var result = JSON.parse(spRequest.responseText);
        reactHandler.setState({
          listTitles: result.d.results,
          loadingLists: false
        });
      } else if (spRequest.readyState === 4 && spRequest.status !== 200) {
        console.log("Error Occured !");
        reactHandler.setState({
          error: "Error occured"
        });
      }
    };
    spRequest.send();
  }
}
