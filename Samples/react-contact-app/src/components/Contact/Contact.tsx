import * as React from "react";
import styles from "./Contact.module.scss";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
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
          <div className={css("ms-Grid-col ms-xxl4 ms-lg6")} key={key}>
            <div className={css(styles.contactItem, styles.noPrint)}>
              <div className={css("ms-xxl3", styles.contactPhoto)}>
                <img src={item["FileRef"]} />
              </div>
              <div className="ms-Grid-col ms-xxl9">
                <div>{this.getName(item["File"]["Name"])}</div>
                <div>{item["Title"]}</div>
                <div>{item["Abteilung"]}</div>
                <hr />
                <div className="phone">T {item["Telefonnummer"]}</div>
                <div>F {item["Faxnummer"]}</div>
                <div>M {item["Mobil_x0020_gschftl_x002e_"]}</div>
                <div>
                  <a href="mailto:{item['E_x002d_Mail']}">
                    {item["E_x002d_Mail"]}
                  </a>
                </div>
              </div>
            </div>
          </div>
        );
      }
    );
    const printTitles: JSX.Element[] = this.state.listTitles.map(
      (item: string, key: number, listTitles: string[]): JSX.Element => {
        const print = item["Durchwahl"] !== null;
        if (print) {
          return (
            <div className={styles.printDiv} key={key}>
              <div className={styles.onlyPrint}>
                <div>
                  <span className={styles.number}>{item["Durchwahl"]}</span>
                  <span className={styles.name}>
                    {this.getName(item["File"]["Name"])}
                  </span>
                  <span className={styles.department}>{item["Abteilung"]}</span>
                </div>
              </div>
            </div>
          );
        }
      }
    );
    return (
      <div className={css("ms-Grid", styles.contact)}>
        <div className={styles.container}>
          <div className={css("ms-Grid-row", styles.row)}>
            <div className="ms-Grid-col ms-xxl12 ms-lg12">
              <header className={styles.noPrint}>
                {" "}
                <h1 className={styles.header}>{this.props.description} </h1>
                <input
                  className={styles.searchBox}
                  type="text"
                  ref="filterInput"
                  placeholder="Suche..."
                  onChange={this.filterUpdate}
                />
                <i
                  onClick={this.print}
                  className={css(
                    "ms-Icon ms-Icon--Print x-hidden-focus",
                    styles.printIcon
                  )}
                  aria-hidden="true"
                />
              </header>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div>
              {this.state.loadingLists && <span>Loading lists...</span>}
              {this.state.error && (
                <span>
                  An error has occurred while loading contact:{" "}
                  {this.state.error}
                </span>
              )}
              {this.state.error === null &&
                titles && (
                  <div>
                    <div className={styles.contactListItem}>{titles}</div>
                    <div className={styles.contactListItemPrint}>
                      {/* <div className={styles.printDiv}>
                        <div className={styles.onlyPrint}>
                          <div>
                            <span className={styles.number}>Nr.</span>
                            <span className={styles.name}>Name</span>
                            <span className={styles.department}>Abteilung</span>
                          </div>
                        </div>
                      </div> */}
                      {printTitles}
                    </div>
                  </div>
                )}
            </div>
          </div>
        </div>
      </div>
    );
  }

  filterUpdate(event) {
    var searchText = event.target.value.toLowerCase();
    var filteredItems = this.state.allItems;
    filteredItems = filteredItems.filter(function(item) {
      return (
        item["File"]["Name"].toLowerCase().startsWith(searchText) ||
        item["Abteilung"].toLowerCase().startsWith(searchText) ||
        (item["Durchwahl"] !== null &&
          item["Durchwahl"].toLowerCase().startsWith(searchText))
      );
    });
    this.setState({ listTitles: filteredItems });
  }

  print() {
    window.print();
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
      "')/items?$select=File/Name,FileRef,Title,Abteilung,Telefonnummer,Faxnummer,Mobil_x0020_gschftl_x002e_,E_x002d_Mail,Durchwahl&$expand=File&$top=1000";
    spRequest.open("GET", listUrl, true);
    spRequest.setRequestHeader("Accept", "application/json;odata=verbose");

    spRequest.onreadystatechange = function() {
      if (spRequest.readyState === 4 && spRequest.status === 200) {
        var result = JSON.parse(spRequest.responseText);
        var sortedResults = result.d.results.sort(function(a, b) {
          return a.File.Name > b.File.Name
            ? 1
            : b.File.Name > a.File.Name
              ? -1
              : 0;
        });
        reactHandler.setState({
          listTitles: sortedResults,
          allItems: sortedResults,
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
