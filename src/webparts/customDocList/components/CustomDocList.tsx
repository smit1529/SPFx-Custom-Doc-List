import * as React from 'react';
import styles from './CustomDocList.module.scss';
import { ICustomDocListProps } from './ICustomDocListProps';
import { ColorClassNames } from '@uifabric/styling/lib';

export interface IDataCol {
  dataValues: IData[];
}

export interface IData {
  ID: string;
  Title: string;
}

export default class CustomDocList extends React.Component<ICustomDocListProps, IDataCol> {
  public constructor(props: ICustomDocListProps, state: IDataCol) {
    super(props);
    this.state = {
      dataValues: [] as IData[]
    };
  }

  public componentDidMount() {
    if (this.props.ListName != undefined)
      this.getListData();
  }

  public componentDidUpdate(prevProps) {
    if (prevProps["ListName"] != this.props.ListName || prevProps["ColumnsToDisplay"] != this.props.ColumnsToDisplay || prevProps["SortByField"] != this.props.SortByField || prevProps["MaxResultsProp"] != this.props.MaxResultsProp) {
      this.getListData();
    }
  }

  public render(): React.ReactElement<ICustomDocListProps> {
    var rh = this;
    return (
      <div className={styles.panelStyle}>
        <br></br>
        <div className={styles.tableCaptionStyle}>Demo: Retrieve SPList Items using Dynamic Web-part</div>
        <br></br>
        <div className={styles.tableStyle}>
          {rh.props.WebPartTitle != undefined && rh.props.ListName != undefined
            ? <div className={styles.rowCaptionStyle}>{rh.props.WebPartTitle + " - " + rh.props.ListName}</div>
            : ""}
          <div className={styles.rowStyle}>
            {rh.props.ColumnsToDisplay != undefined
              ? rh.props.ColumnsToDisplay.map(function (item, key) {
                return (
                  <div className={styles.headerStyle}>
                    {rh.props.ListColumns.filter(x => x.InternalName == item)[0]["Title"]}
                  </div>
                );
              })
              : ""}
          </div>
          {rh.state.dataValues.map(function (item, key) {
            return (
              <div className={styles.rowStyle}>
                {rh.props.ColumnsToDisplay.map(function (i, v) {
                  var col = rh.props.ListColumns.filter(x => x.InternalName == i)[0];
                  if (col.TypeDisplayName == "Single line of text") {
                    return (
                      <div className={styles.CellStyle}>
                        <span>{item[i]}</span>
                      </div>
                    );
                  }
                  else if (col.TypeDisplayName == "Multiple lines of text") {
                    return (
                      <div className={styles.CellStyle}>
                        <span dangerouslySetInnerHTML={{ __html: item[i] }}></span>
                      </div>
                    );
                  }
                  else if (col.TypeDisplayName == "Hyperlink or Picture") {
                    return (
                      <div className={styles.CellStyle}>
                        <a href={item[i] != null ? item[i]["Url"] : "#"}>{item[i] != null ? item[i]["Description"] : ""}</a>
                      </div>
                    );
                  }
                  else if (col.TypeDisplayName == "Attachments") {
                    if (item[i]) {
                      return (
                        <div className={styles.CellStyle}>
                          <a href={item["AttachmentFiles"][0]["ServerRelativeUrl"] + "?web=1"} target="_blank">
                            {item["AttachmentFiles"][0]["FileName"]}
                          </a>
                        </div>
                      );
                    }
                    else {
                      return (
                        <div className={styles.CellStyle}>
                          <span>No Attachment Found</span>
                        </div>
                      );
                    }
                  }
                  else if (col.TypeDisplayName == "Number") {
                    return (
                      <div className={styles.CellStyle}>
                        <span>{item[i]}</span>
                      </div>
                    );
                  }
                  else if (col.TypeDisplayName == "Yes/No") {
                    return (
                      <div className={styles.CellStyle}>
                        <span>{item[i] == true ? "Yes" : "No"}</span>
                      </div>
                    );
                  }
                  else {
                    // console.log(col.Title + " - " + col.InternalName + " - " + col.TypeDisplayName)
                    return (
                      <div className={styles.CellStyle}>
                        <span>Else</span>
                      </div>
                    );
                  }
                })}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  public getListData() {
    var rh = this;
    // console.log(rh.props.SortByField, rh.props.ListColumns);
    var lstName = this.props.ListName;
    var siteUrl = this.props.siteAbsoluteUrl;
    var selectedColumns = this.props.ColumnsToDisplay != undefined ? this.props.ColumnsToDisplay.join(',') : "*";
    var sortByColumn = this.props.SortByField != undefined && this.props.SortByField != "None" ?
      rh.props.ListColumns.filter(x => x.Title == rh.props.SortByField)[0]["InternalName"]
      : "";

    var reactHandler = this;
    var url = siteUrl + `/_api/web/lists/getbytitle('` + lstName + `')/items?$select=AttachmentFiles,` + selectedColumns + `&$expand=AttachmentFiles&$orderby=` + sortByColumn + `&$top=` + this.props.MaxResultsProp + ``;
    // console.log(url);

    fetch(url, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' }
    })
      .then((res) => res.json())
      .then(
        (result) => {
          reactHandler.setState({ dataValues: result.value });
        },
        (error) => {
          console.log(error);
        }
      );
  }
}
