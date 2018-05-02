import * as React from "react";
import styles from "./Contact.module.scss";
import { css } from "office-ui-fabric-react";
import { IContactProps } from "./IContactProps";
import { IContactState } from "./IContactState";
require("sp-init");
require("microsoft-ajax");
require("sp-runtime");
require("sharepoint");

export class Contact extends React.Component<IContactProps, IContactState> {
  constructor(props?: IContactProps, context?: any) {
    super();
    this.state = {
      listTitles: [],
      contactList: [],
      allItems: [],
      checkbox: false,
      filterText: "",
      loadingLists: false,
      error: null
    };

    this.getMoreContact = this.getMoreContact.bind(this);
    this.filterUpdate = this.filterUpdate.bind(this);
  }
  // componentWillMount = () => {
  //   this.selectedCheckboxes = new Set();
  // };

  componentDidMount() {
    this.getContactListRest();
    // this.getContactListCsom();
  }

  public render(): React.ReactElement<IContactProps> {
    const titles: JSX.Element[] = this.state.listTitles.map(
      (item: string, key: number, listTitles: string[]): JSX.Element => {
        return (
          <div className={styles.contactItem} key={key}>
            <div className={styles.contactPhoto + " " + styles.noPrint}>
              <img src={item["FileRef"]} />
            </div>
            <div>
              <h1>{this.getName(item["File"]["Name"])}</h1>
              <span className={styles.noPrint}>{item["Title"]}</span>
              <span className={styles.noPrint}>{item["Abteilung"]}</span>
            </div>
            <div>
              <div className="phone">T {item["Telefonnummer"]}</div>
              <div className={styles.noPrint}>F {item["Faxnummer"]}</div>
              <div className={styles.noPrint}>
                M {item["Mobil_x0020_gschftl_x002e_"]}
              </div>
              <div className={styles.noPrint}>
                <a href="mailto:{item['E_x002d_Mail']}">
                  {item["E_x002d_Mail"]}
                </a>
              </div>
            </div>
          </div>
        );
      }
    );
    return (
      <div className={styles.contact}>
        <div className={styles.container}>
          <div
            className={css(
              "ms-Grid-row ms-bgColor-teal ms-fontColor-white",
              styles.row
            )}
          >
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <h1 className={styles.header}>{this.props.description} </h1>
              <header>
                <input
                  type="text"
                  ref="filterInput"
                  placeholder="Type to filter.."
                  onChange={this.filterUpdate}
                />
                <label>
                  Is going:
                  <input
                    name="isGoing"
                    type="checkbox"
                    checked={this.state.checkbox}
                    onChange={this.filterUpdate}
                  />
                </label>
              </header>
              <div className="contact-list-container">
                {this.state.loadingLists && <span>Loading lists...</span>}
                {this.state.error && (
                  <span>
                    An error has occurred while loading contact:{" "}
                    {this.state.error}
                  </span>
                )}
                {this.state.error === null &&
                  titles && (
                    <div className={styles.contactListItem}>{titles}</div>
                  )}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  filterUpdate(event) {
    if (event.target.type === "checkbox") {
      alert("Test");
    }
    var filteredItems = this.state.allItems;
    filteredItems = filteredItems.filter(function(item) {
      return (
        item["File"]["Name"]
          .toLowerCase()
          .search(event.target.value.toLowerCase()) !== -1
      );
    });
    this.setState({ listTitles: filteredItems });
  }

  private getName(docName) {
    return docName.substring(0, docName.indexOf("."));
  }

  private getContactListRest() {
    var reactHandler = this;
    reactHandler.setState({
      loadingLists: true,
      listTitles: [],
      allItems: [],
      error: null
    });

    var spRequest = new XMLHttpRequest();
    var currentMonth = new Date().getMonth() + 1;
    var listUrl =
      this.props.siteUrl +
      "/_api/web/lists/GetByTitle('" +
      this.props.listName +
      "')/items?$select=File/Name,FileRef,Title,Abteilung,Telefonnummer,Faxnummer,Mobil_x0020_gschftl_x002e_,E_x002d_Mail&$expand=File&$top=1000";
    spRequest.open("GET", listUrl, true);
    spRequest.setRequestHeader("Accept", "application/json;odata=verbose");

    spRequest.onreadystatechange = function() {
      if (spRequest.readyState === 4 && spRequest.status === 200) {
        var result = JSON.parse(spRequest.responseText);
        reactHandler.setState({
          listTitles: result.d.results,
          allItems: result.d.results,
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

  private getMoreContact(): void {
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
