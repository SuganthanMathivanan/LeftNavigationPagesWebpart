//import jQuery from 'jQuery';
import * as React from 'react';
import styles from './LeftNavigationMenuPagesWebpart.module.scss';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import {
  List,
  PrimaryButton
} from 'office-ui-fabric-react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';

import {
  css,
  DocumentCard,
  DocumentCardLocation,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity,
  Spinner
} from 'office-ui-fabric-react';
import * as Enumerable from 'linq';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import * as pnp from 'sp-pnp-js';
import Web, { ODataDefaultParser } from 'sp-pnp-js';
import * as $ from 'jquery';


export interface IDisplayDocumentsProps {
  context: WebPartContext;
  description: string;

}
export interface IMyRecentItems {
  Id?: number;
  Title?: string;
  ItemUrl?: string;
  Description?: string;
}
export interface IMyRecentsTopBarState {

  myRecentItems: IMyRecentItems[];
  hideDialog: boolean;

}
let _items: IDocument[] = [];
let _favorites: IDocument[] = [];

export interface IDetailsListDocumentsExampleState {
  columns: IColumn[];
  favoriteColumns: IColumn[];
  items: IDocument[];
  favorites: IDocument[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  myRecentItems: IMyRecentItems[];
  hideDialog: boolean;
}

export interface IDocument {
  [key: string]: any;
  name: String;
  value: String;
  iconName: string;
  type: any;
  clientRef: any;
  clientName: any,
  href: any;
}
let searchResultsLgth: number = 2;
var currFavID = 0;
var thisDuplicating;
var absUrlSubStr;
let seeMoreUrl: string = "/Search/Pages/results.aspx?k=contentclass:STS_List_DocumentLibrary%20Title:";
var selectionCount = 0;
export default class DisplayDocuments extends React.Component<IDisplayDocumentsProps, IDetailsListDocumentsExampleState> {
  private _context: WebPartContext;
  private _currentWebUrl: string;
  private _selection: Selection;
  constructor(props: IDisplayDocumentsProps) {
    super(props);
    this._context = props.context;
    this._currentWebUrl = this._context.pageContext.web.absoluteUrl;
    this.getClientRecents();
    this.getClientFavourites('new');

    const _columns: IColumn[] = [
      {
        key: 'column1',
        name: 'File Type',
        headerClassName: 'DetailsListExample-header--FileIcon',
        className: 'DetailsListExample-cell--FileIcon',
        iconClassName: 'DetailsListExample-Header-FileTypeIcon',
        ariaLabel: 'Column operations for File type',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <img src={item.iconName} className={'DetailsListExample-documentIconImage'} />;
        }
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 150,
        maxWidth: 150,
        isRowHeader: true,
        //isResizable: true,
        //isSorted: true,
        //isSortedDescending: false,
        //sortAscendingAriaLabel: 'Sorted A to Z',
        //sortDescendingAriaLabel: 'Sorted Z to A',
        //onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true,
        onRender: (item: IDocument) => {
          return <a href={item.href} className={styles.titleAnchorStyle}>{item.name}</a>;
        }
      },
      {
        key: 'column3',
        name: 'Client Reference',
        fieldName: 'clientreference',
        minWidth: 130,
        maxWidth: 130,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.clientRef}</span>;

        },
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Client Name',
        fieldName: 'clientname',
        minWidth: 130,
        maxWidth: 130,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.clientName}</span>;

        },
        isPadded: true
      }
    ];
    const _favoriteColumns: IColumn[] = [
      {
        key: 'column1',
        name: 'File Type',
        headerClassName: 'DetailsListExample-header--FileIcon',
        className: 'DetailsListExample-cell--FileIcon',
        iconClassName: 'DetailsListExample-Header-FileTypeIcon',
        ariaLabel: 'Column operations for File type',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <img src={item.iconName} className={'DetailsListExample-documentIconImage'} />;
        }
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 250,
        maxWidth: 250,
        isRowHeader: true,
        //isResizable: true,
        //isSorted: true,
        //isSortedDescending: false,
        //sortAscendingAriaLabel: 'Sorted A to Z',
        //sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true,
        onRender: (item: IDocument) => {
          return <a href={item.href} className={styles.titleAnchorStyle}>{item.name}</a>;
        }
      },
      {
        key: 'column3',
        name: 'Client Reference',
        fieldName: 'clientreference',
        minWidth: 130,
        maxWidth: 130,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.clientRef}</span>;

        },
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Client Name',
        fieldName: 'clientname',
        minWidth: 130,
        maxWidth: 130,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.clientName}</span>;

        },
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Actions',
        fieldName: 'type',
        minWidth: 100,
        maxWidth: 100,
        className: styles.actionsCol,
        //isResizable: true,
        //isCollapsable: true,
        //data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <div dangerouslySetInnerHTML={{ __html: item.type }} />;
        },
        isPadded: true
      }
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
          isModalSelection: this._selection.isModal()
        });
      }
    });

    this.state = {
      //items: _items,
      items: [],
      favorites: [],
      columns: _columns,
      favoriteColumns: _favoriteColumns,
      selectionDetails: this._getSelectionDetails(),
      //isModalSelection: this._selection.isModal(),
      isModalSelection: this._selection.isModal(),
      isCompactMode: false,
      myRecentItems: [],
      hideDialog: true
    };

  }
  private searchboxval: string;
  public render(): React.ReactElement<IDisplayDocumentsProps> {
    const { columns, isCompactMode, items, selectionDetails, favorites, favoriteColumns } = this.state;
    const displayStyle = {
      display: 'none',
    };
    thisDuplicating = this;

    return (
      <div className={styles.displayDocuments}>
        <div className={styles.container}>
          <div className={styles.row}>
            {/*<div className={'ms-Grid '+ styles.column }>-->*/}
            <div>
              <Dialog
                hidden={this.state.hideDialog}
                onDismiss={this._closeDialog}
                dialogContentProps={{
                  type: DialogType.normal,
                  title: 'Important!',
                  subText:
                    "Are you sure want to delete this favourite? Click 'Yes' to continue else click 'No' to cancel."
                }}
                modalProps={{
                  isBlocking: true,
                  containerClassName: 'ms-dialogMainOverride'
                }}
              >
                <DialogFooter>
                  <PrimaryButton onClick={this._deleteFav} text="Yes" />
                  <DefaultButton onClick={this._closeDialog} text="No" />
                </DialogFooter>
              </Dialog>
            </div>
            <div className={'ms-Grid'} dir="ltr">
              <div className="search ms-Grid-row">
                <div className={'searchBocContains ' + styles.searchBocContains}>
                  <div className="ms-SearchBoxExample">
                    <div className={'ms-Grid-col ms-lg12 ' + styles.searchboxtitle}>Clients</div>
                    <div className={'ms-Grid-col ms-lg12 ' + styles.margin}>
                      <SearchBox
                        className={'ms-Grid-col ms-sm8 ms-lg9 ' + styles.searchbox}
                        placeholder="Search"
                        onSearch={newvalue => this.clientsearchresults(newvalue)}
                        onFocus={() => console.log('onFocus called')}
                        onBlur={onblurval => this.onBlurSearch(onblurval)}
                        onChange={onchangeVal => this.onChangeSearch(onchangeVal)}
                      />
                      <PrimaryButton className={'ms-Grid-col ms-sm3 ms-lg2 ' + styles.button} text="Search" onClick={() => { this.clientsearchresults(this.searchboxval); }} />
                      <div style={displayStyle} className={'ms-Grid-col ms-lg5 searchResults ' + styles.searchResults}>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <div className={'ms-Grid-row recentDocs ' + styles.recent}>
                <div className={'ms-Grid-col ms-lg12 accordion1 ' + styles.accordion1}><p className={'ms-Grid-col ms-lg11 accordiontitle1 ' + styles.accordionTitle} >Recent Clients</p><div className={styles.accordionicon4}><i className={'ms-Grid-col ms-lg1 accordionicon1 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon1} aria-hidden="true"></i></div></div>
                <div className={'containsRecents ' + styles.panel}>
                  <MarqueeSelection selection={this._selection}>
                    <DetailsList
                      items={items}
                      //compact={isCompactMode}
                      columns={columns}
                      selectionMode={this.state.isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                      //setKey="set"
                      layoutMode={DetailsListLayoutMode.justified}
                      isHeaderVisible={true}
                      selection={this._selection}
                    //selectionPreservedOnEmptyClick={true}
                    //onItemInvoked={this._onItemInvoked}
                    //enterModalSelectionOnTouch={true}
                    />
                  </MarqueeSelection>
                </div>
              </div>
              <div className={'ms-Grid-row ClientFavourites ' + styles.ClientFavourites}>
                <div className={'ms-Grid-col ms-lg12 accordion2 ' + styles.accordion2}><p className={'ms-Grid-col ms-lg11 accordiontitle2 ' + styles.accordionTitle} >Favourite Clients</p><div className={styles.accordionicon4}><i className={'ms-Grid-col ms-lg1 accordionicon2 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon2} aria-hidden="true"></i></div></div>
                <div className={'containsFav ' + styles.panel}>
                  <MarqueeSelection selection={this._selection}>
                    <DetailsList
                      items={this.state.favorites}
                      compact={isCompactMode}
                      columns={favoriteColumns}
                      selectionMode={this.state.isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                      //setKey="set"
                      layoutMode={DetailsListLayoutMode.justified}
                      isHeaderVisible={true}
                    //selection={this._selection}
                    //selectionPreservedOnEmptyClick={true}
                    //onItemInvoked={this._onItemInvoked}
                    //enterModalSelectionOnTouch={true}
                    />
                  </MarqueeSelection>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  };
  private _deleteFav = (): void => {
    pnp.sp.site.rootWeb.lists.getByTitle('Favourite').items.getById(currFavID).delete()
      .then(res => {
        console.log('Happy');

        $(".favoritebox" + currFavID).remove();
        this.setState({ hideDialog: true });
        this.getClientFavourites('again');
      });
  };
  public onBlurSearch(onblurval) {
    console.log(onblurval);
    if (onblurval.path[0].value == '') {
      $('.searchResults').css('display', 'none');
    }
  }
  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  };
  public onChangeSearch(onchangeText: string): void {
    this.searchboxval = onchangeText;
    if (onchangeText == "") {
      $('.searchResults').html('');
      $('.searchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
      $('.searchResults').removeAttr('style');
    }

  }
  public async clientsearchresults(text: string) {
    if (text != undefined && text != "") {
      $('.searchResults').html('');
      let absUrl: string = this._context.pageContext.web.absoluteUrl + "/Clients";
      var absurlIndex = absUrl.indexOf('.com');
      absUrlSubStr = absUrl.substring(0, absurlIndex + 4);
      var webAbs = this._context.pageContext.web.absoluteUrl;
      var icon = absUrlSubStr + "/_layouts/15/images/itdl.png?rev=44";
      /*pnp.sp.site.getDocumentLibraries(absUrl)
        .then((data) => {
          var strTitle = "";
          let flag: number = 0;
          $.each(data, function (index, value) {
            var ind;
            let title: String = value.Title.toLowerCase();
            //ind = title.indexOf(text.toLowerCase());
            ind = title.startsWith(text.toLowerCase());
            if (ind != false) {

              if (flag < searchResultsLgth) {
                strTitle += '<div class=' + styles.searchResHover + '><div><div class=' + styles.folderIconContain + '><img src=' + icon + ' className="DetailsListExample-documentIconImage" /></div><a class=' + styles.anchorlib + ' href=' + encodeURI(value.AbsoluteUrl) + '>' + value.Title + '</a></div></div>';
              }
              flag++;
            }
          });
          if (flag == 0) {
            $('.searchResults').html('');
            $('.searchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
            $('.searchResults').removeAttr('style');
          }
          else if (flag > searchResultsLgth) {
            var href = webAbs + seeMoreUrl + text + "*";
            strTitle += '<div class=' + styles.seemore + '><a class=' + styles.seemorelink + ' href=' + href + '>see more</a></div>';
            $('.searchResults').append(strTitle);
            $('.searchResults').removeAttr('style');
          }
          else {
            $('.searchResults').append(strTitle);
            $('.searchResults').removeAttr('style');
          }
        }).catch(function (err) {
          alert(err);
        });*/
      $.ajax({
        //url: absurl + "/_api/search/query?querytext=%27(Path:" + absurl + ")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20ContentType:Document%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%20NOT(FileExtension:mht%20OR%20FileExtension:aspx%20OR%20FileExtension:html%20OR%20FileExtension:htm)%27&rowlimit=50&bypassresulttypes=false&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,SPWebUrl,PTLClientReference%27&sortlist=%27LastModifiedTime:descending%27&enablesorting=true&refiners='PTLClientReference'",
        //url: absurl + "_api/search/query?querytext=%27(Path:" + absurl + ")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20ContentType:Managed%20Document%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%27&rowlimit=50&bypassresulttypes=false&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,SPWebUrl,RefinableString10,PTLClientReference%27&sortlist=%27LastModifiedTime:descending%27&enablesorting=true&refiners='RefinableString10'",
        //url: absUrl + "/_api/search/query?querytext=%27(PTLClientReference:" + text + "* OR PTLClientName:" + text + "*)%27&properties=%27SourceName:MyRecentDocsScopeWeb,SourceLevel:SPSite%27&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,RefinableString10,PTLClientReference,PTLClientName,SPWebUrl%27&refiners=%27RefinableString10%27&sortlist=%27LastModifiedTime:descending%27&rowlimit=500",
        url: absUrl + "/_api/search/query?querytext=%27(PTLClientReference:" + text + "* OR PTLClientName:" + text + "*)%27&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,RefinableString10,PTLClientReference,PTLClientName,SPWebUrl%27&refiners=%27RefinableString10%27&sortlist=%27LastModifiedTime:descending%27&rowlimit=500",
        method: "GET",
        headers: {
          "accept": "application/json;odata=verbose",   //It defines the Data format
        },
        cache: false,
        success: function (data) {
          var strTitle = "";
          let flag: number = 0;
          if (data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.length != 0) {
            $.each(data.d.query.PrimaryQueryResult.RefinementResults.Refiners.results[0].Entries.results, function (index, value) {
              var ind;
              //let title: String = value.RefinementName.toLowerCase();
              let title: String = value.RefinementName;
              //ind = title.indexOf(text.toLowerCase());
              var libUrl = absUrl + "/" + title;
              //ind = title.startsWith(text.toLowerCase());
              //if (ind != false) {

                if (flag < searchResultsLgth) {
                  strTitle += '<div class=' + styles.searchResHover + '><div><div class=' + styles.folderIconContain + '><img src=' + icon + ' className="DetailsListExample-documentIconImage" /></div><a class=' + styles.anchorlib + ' href=' + encodeURI(libUrl) + '>' + title + '</a></div></div>';
                }
                flag++;
              //}
            });
          }
          else {
            $('.searchMatterResults').html('');
            $('.searchMatterResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
            $('.searchMatterResults').removeAttr('style');
          }
          if (flag == 0) {
            $('.searchResults').html('');
            $('.searchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
            $('.searchResults').removeAttr('style');
          }
          else if (flag > searchResultsLgth) {
            var href = webAbs + seeMoreUrl + text + "*";
            strTitle += '<div class=' + styles.seemore + '><a class=' + styles.seemorelink + ' href=' + href + '>see more</a></div>';
            $('.searchResults').append(strTitle);
            $('.searchResults').removeAttr('style');
          }
          else {
            $('.searchResults').append(strTitle);
            $('.searchResults').removeAttr('style');
          }


        },
        error: function (data) {
          console.log(data);
        }
      });
    }
  }

  private _getSelectionDetails(): string {
    selectionCount = 0;

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as any).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    let newItems: IDocument[] = items.slice();
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter((currCol: IColumn, idx: number) => {
      return column.key === currCol.key;
    })[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    newItems = this._sortItems(newItems, currColumn.fieldName || '', currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  };

  private _sortItems = (items: IDocument[], sortBy: string, descending = false): IDocument[] => {
    if (descending) {
      return items.sort((a: IDocument, b: IDocument) => {
        if (a[sortBy] < b[sortBy]) {
          return 1;
        }
        if (a[sortBy] > b[sortBy]) {
          return -1;
        }
        return 0;
      });
    } else {
      return items.sort((a: IDocument, b: IDocument) => {
        if (a[sortBy] < b[sortBy]) {
          return -1;
        }
        if (a[sortBy] > b[sortBy]) {
          return 1;
        }
        return 0;
      });
    }
  };
  public async getClientRecents() {
    $('.ms-SearchBox-iconContainer').css("color", "black")
    var htmlrecentstr = '';
    var refinementname;
    var absurl = this._context.pageContext.web.absoluteUrl + "/Clients/";
    var absurlIndex = absurl.indexOf('.com');
    absUrlSubStr = absurl.substring(0, absurlIndex + 4);
    _items = [];
    $.ajax({
      //url: absurl + "/_api/search/query?querytext=%27(Path:" + absurl + ")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20ContentType:Document%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%20NOT(FileExtension:mht%20OR%20FileExtension:aspx%20OR%20FileExtension:html%20OR%20FileExtension:htm)%27&rowlimit=50&bypassresulttypes=false&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,SPWebUrl,PTLClientReference%27&sortlist=%27LastModifiedTime:descending%27&enablesorting=true&refiners='PTLClientReference'",
      //url: absurl + "_api/search/query?querytext=%27(Path:" + absurl + ")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20ContentType:Managed%20Document%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%27&rowlimit=50&bypassresulttypes=false&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,SPWebUrl,RefinableString10,PTLClientReference%27&sortlist=%27LastModifiedTime:descending%27&enablesorting=true&refiners='RefinableString10'",
      url: absurl + "_api/search/query?querytext=%27*%27&properties=%27SourceName:MyRecentDocsScopeWeb,SourceLevel:SPSite%27&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,RefinableString10,PTLClientReference,PTLClientName,SPWebUrl%27&refiners=%27RefinableString10%27&sortlist=%27LastModifiedTime:descending%27&rowlimit=500",
      method: "GET",
      headers: {
        "accept": "application/json;odata=verbose",   //It defines the Data format
      },
      cache: false,
      success: function (clientrecents) {
        let flag: number = 0;
        //var p = new Object(clientrecents).toString();
        //console.log(JSON.parse(clientrecents.d));
        var relevantResults = clientrecents.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
        if (clientrecents.d.query.PrimaryQueryResult.RefinementResults != null) {
          $.each(clientrecents.d.query.PrimaryQueryResult.RefinementResults.Refiners.results[0].Entries.results, function (index, value) {
            var hrefUrl = absurl + value.RefinementName;
            refinementname = value.RefinementName;
            findValues(clientrecents.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results);
            console.log(clientVal + "---" + clientName);
            if (flag < 20) {
              flag++;
              _items.push({
                name: value.RefinementName,
                value: value.RefinementName,
                iconName: absUrlSubStr + "/_layouts/15/images/itdl.png?rev=44",
                type: "Document Library",
                href: hrefUrl,
                clientRef: clientVal,
                clientName: clientName,
              });
            }
          });
        }
        thisDuplicating.setState({ items: _items });

        var accordion1 = document.getElementsByClassName("accordion1");
        var i;
        for (i = 0; i < accordion1.length; i++) {
          accordion1[i].addEventListener("click", function () {
            this.classList.toggle("active");
            var panel = this.nextElementSibling;
            if (panel.style.maxHeight) {
              panel.style.maxHeight = null;
              $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
              $('.accordiontitle1').css("color", "rgba(106, 191, 52, 1)");
              $('.accordionicon1').css("color", "gray");
              $('.accordionicon1').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
            } else {
              panel.style.maxHeight = panel.scrollHeight + "px";
              $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
              $('.accordiontitle1, .accordionicon1').css("color", "white");
              $('.accordionicon1').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
            }
          });
        }
        $('.accordion1').trigger('click');

      },
      error: function (data) {
        console.log(data);
      }
    });
    var clientVal;
    var clientName;
    function findValues(harddata) {
      var data = harddata;
      var queryResult = Enumerable.from(harddata)
        .where(function (x, y) { 
          console.log(x);
          return data[y].Cells.results; 
         })
        //.select("$.Cells.results.value+':'+"+refinementName+"")
        .select(
          function (x, y) {
            console.log(data[y].Cells.results);
            var key = data[y].Cells.results.find(x => x.Key === 'PTLClientReference').Value;
            if (key == refinementname) {
              clientVal = data[y].Cells.results.find(x => x.Key === 'PTLClientReference').Value;

              clientName = data[y].Cells.results.find(x => x.Key === 'PTLClientName').Value;
              return clientVal + "@#@" + clientName;
            }
          })
        .toArray();
      console.log(queryResult);
      return queryResult;
    }
  }

  public async getClientFavourites(para) {
    var htmlStr = "";
    //this.setState({ favorites: [] });
    _favorites = [];
    pnp.sp.web.currentUser.get().then(result => {
      let currUserId: number = result.Id;
      let queryStr: string = "Author eq " + currUserId + "and PTLCategory eq 'Client'";
      pnp.sp.web.lists.getByTitle("Favourite")
        .items
        .select(
          "Id",
          "Title",
          "PTLItemURL",
          "PTLFavouriteMetadata",
          "PTLCategory"
        )
        .filter(queryStr)
        .get()
        .then((data) => {
          //console.log('happy');
          selectionCount = data.length;
          this._getSelectionDetails();
          $.each(data, function (index, value) {
            //htmlStr += '<div id="favoritebox' + index + '" class="ms-Grid-col ms-lg5 favoritebox' + value.ID + ' ' + styles.favoritebox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg6 clientAnchorlib"><div class="clientAnchorlib ' + styles.clientAnchorlib + '">' + value.Title + '</div></div><div title="Launch Favorite" class="drillbox ms-Grid-col ms-lg2 drillbox ' + styles.launchbox + '"><a href="' + value.PTLItemURL + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></a></div><div title="Delete Favorite" class="deletebox ms-Grid-col ms-lg2 ' + styles.deletebox + '" value = "' + value.ID + '"><i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></div></div>';
            var metadata = value.PTLFavouriteMetadata.split('#@#');

            _favorites.push({
              name: value.Title,
              value: value.Title,
              iconName: absUrlSubStr + "/_layouts/15/images/itdl.png?rev=44",
              type: '<a class="link" data-value="' + value.PTLItemURL + '"><div title="Launch Favourite" class="drillbox ms-Grid-col ms-lg2 drillbox ' + styles.launchbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></a><div title="Delete Favourite" class="deletebox ms-Grid-col ms-lg2 ' + styles.deletebox + '" value = "' + value.ID + '"><i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></div>',
              href: value.PTLItemURL,
              /*dateModified: "",
              dateModifiedValue: 0,
              fileSize: '25',
              fileSizeRaw: 30,
              displayName: "Document Library",*/
              clientRef: metadata[1],
              clientName: metadata[0],
            });
          });
          thisDuplicating.setState({ favorites: _favorites });
          //$('.containsFav').append(htmlStr);
          /*if(para == 'again'){
            this.setState({ favorites: _favorites });
          }*/
          this.setState({ hideDialog: true });
          if (para == 'new') {
            $(".deletebox").click(function (e) {
              e.preventDefault();
              currFavID = parseInt($(this).attr("value"));
              thisDuplicating._showDialog();
              return false;

            });
            $(".link").click(function (e) {
              e.preventDefault();
              var val = $(this).attr("data-value");
              //thisDuplicating._showDialog();
              window.location.assign(val);
              return false;

            });

            var accordion2 = document.getElementsByClassName("accordion2");
            var i;
            for (i = 0; i < accordion2.length; i++) {
              accordion2[i].addEventListener("click", function () {

                this.classList.toggle("active");
                var panel = this.nextElementSibling;
                if (panel.style.maxHeight) {
                  panel.style.maxHeight = null;

                  $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
                  $('.accordiontitle2').css("color", "rgba(106, 191, 52, 1)");
                  $('.accordionicon2').css("color", "gray");
                  $('.accordionicon2').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
                } else {
                  panel.style.maxHeight = panel.scrollHeight + "px";
                  //$(this).css("background-color", "rgba(55, 55, 55, 1)");
                  //$(this).css("cssText", "border:none !important");
                  $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
                  $('.accordiontitle2, .accordionicon2').css("color", "white");
                  $('.accordionicon2').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
                }
              }, false);
            }
          }

        });
    });
  }
}

