import * as React from "react";
import styles from "./Birthday.module.scss";
import { css } from "office-ui-fabric-react";
import { IBirthdayProps } from "./IBirthdayProps";
import { IBirthdayState } from "./IBirthdayState";
require("sp-init");
require("microsoft-ajax");
require("sp-runtime");
require("sharepoint");

export class Birthday extends React.Component<IBirthdayProps, IBirthdayState> {
  constructor(props?: IBirthdayProps, context?: any) {
    super();
    this.state = {
      listTitles: [],
      birthdayList: [],
      loadingLists: false,
      error: null
    };

    this.getMoreBirthday = this.getMoreBirthday.bind(this);
  }

  componentDidMount() {
    this.getBirthdayListRest();
    // this.getBirthdayListCsom();
  }

  public render(): React.ReactElement<IBirthdayProps> {
    const titles: JSX.Element[] = this.state.listTitles.map(
      (item: string, key: number, listTitles: string[]): JSX.Element => {
        var birthdayDate = new Date(
          parseInt(item["Anfangszeit"].substring(6, 19))
        ).format("dd.MM.");
        if (birthdayDate == new Date().format("dd.MM.")) {
          return (
            <div
              className={styles.birthdayItem + " " + styles.birthdayToday}
              key={key}
            >
              <p className={styles.head}>
                {birthdayDate} - {item["Titel"]}
              </p>
            </div>
          );
        } else {
          return (
            <div className={styles.birthdayItem} key={key}>
              <p className={styles.head}>
                {birthdayDate} - {item["Titel"]}
              </p>
            </div>
          );
        }
      }
    );
    // const birthday: JSX.Element[] = this.state.birthdayList.map((value: string, key: number, birthdayList: string[]): JSX.Element => {
    //   return <li key={key}>{value}</li>;
    // });
    return (
      <div className={styles.birthday}>
        <div className={styles.container}>
          <div
            className={css(
              "ms-Grid-row ms-bgColor-teal ms-fontColor-white",
              styles.row
            )}
          >
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <h1 className={styles.header}>{this.props.description} </h1>
              <div className="birthday-list-container">
                {this.state.loadingLists && <span>Loading lists...</span>}
                {this.state.error && (
                  <span>
                    An error has occurred while loading birthday:{" "}
                    {this.state.error}
                  </span>
                )}
                {this.state.error === null &&
                  titles && (
                    <div className={styles.birthdayListItem}>{titles}</div>
                  )}
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

  private getBirthdayListRest() {
    var reactHandler = this;
    reactHandler.setState({
      loadingLists: true,
      listTitles: [],
      error: null
    });

    var spRequest = new XMLHttpRequest();
    var currentMonth = new Date().getMonth() + 1;
    var today = new Date();
    var todayMinusFive = new Date();
    todayMinusFive.setDate(today.getDate() - 5);
    var todayPlusFive = new Date();
    todayPlusFive.setDate(today.getDate() + 5);
    var listUrl =
      this.props.siteUrl +
      "/_vti_bin/listdata.svc/" +
      this.props.listName +
      "?$filter=((Anfangszeit ge datetime'" +
      todayMinusFive.toISOString() +
      "') and (Anfangszeit le datetime'" +
      todayPlusFive.toISOString() +
      "'))&$orderby=Anfangszeit asc";
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

  private getBirthdayListCsom(): void {
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

        const birthdayTitles: string[] = [];
        while (listEnumerator.moveNext()) {
          const list: SP.List = listEnumerator.get_current();
          birthdayTitles.push(list.get_title());
        }

        this.setState(
          (
            prevState: IBirthdayState,
            props: IBirthdayProps
          ): IBirthdayState => {
            prevState.birthdayList = birthdayTitles;
            prevState.loadingLists = false;
            return prevState;
          }
        );
      },
      (sender: any, args: SP.ClientRequestFailedEventArgs): void => {
        this.setState({
          loadingLists: false,
          birthdayList: [],
          error: args.get_message()
        });
      }
    );
  }

  private getMoreBirthday(): void {
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
      "')/items";
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
