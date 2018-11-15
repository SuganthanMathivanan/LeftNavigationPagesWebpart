//import jQuery from 'jQuery';
import * as React from 'react';
import styles from './LeftNavigationMenuPagesWebpart.module.scss';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import {
  List,
  PrimaryButton
} from 'office-ui-fabric-react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
//import { IDisplayDocumentsProps } from './IDisplayDocumentsProps';
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
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
//import '*' as Enumerable from 'linq';
import * as Enumerable from 'linq';

import * as pnp from 'sp-pnp-js';
import Web from 'sp-pnp-js';
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
/*export interface IMyRecentsTopBarState {

  myRecentItems: IMyRecentItems[];
  hideDialog: boolean;

}*/
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
  href: any;
  /*dateModified: String;
  dateModifiedValue: number;
  fileSize: String;
  fileSizeRaw: number;
  displayName: string;*/
  clientName: string;
  clientRef: string;
  matterRef: string;
  matterKey: string;
}
let searchResultsLgth: number = 2;
var absUrlSubStr;
let seeMoreUrl: string = "/Search/Pages/results.aspx?k=ContentType:Managed%20Folder%20IsDocument:false%20Title:";
var currFavID = 0;
var thisDuplicating;
export default class MatterDocuments extends React.Component<IDisplayDocumentsProps, IDetailsListDocumentsExampleState> {
  private _context: WebPartContext;
  private _currentWebUrl: string;
  private _selection: Selection;
  constructor(props: IDisplayDocumentsProps) {
    super(props);
    this._context = props.context;
    this._currentWebUrl = this._context.pageContext.web.absoluteUrl;
    /*this.state = {
        myRecentItems:[],
        hideDialog: true
    };*/
    this.getMatterRecents();
    this.getMatterFavourites('new');
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
          return <div dangerouslySetInnerHTML={{ __html: item.iconName }} />;
        }
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 150,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
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
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
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
        key: 'column4',
        name: 'Client Name',
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
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
        key: 'column5',
        name: 'Matter Reference',
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.matterRef}</span>;

        },
        isPadded: true
      },
      {
        key: 'column6',
        name: 'Matter Key Desc',
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.matterKey}</span>;

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
          return <div dangerouslySetInnerHTML={{ __html: item.iconName }} />;
        }
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 130,
        maxWidth: 130,
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
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
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
        key: 'column4',
        name: 'Client Name',
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
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
        key: 'column5',
        name: 'Matter Reference',
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.matterRef}</span>;

        },
        isPadded: true
      },
      {
        key: 'column6',
        name: 'Matter Key Desc',
        fieldName: 'type',
        minWidth: 70,
        maxWidth: 70,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span className={styles.authortext}>{item.matterKey}</span>;

        },
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Actions',
        fieldName: 'type',
        minWidth: 50,
        maxWidth: 50,
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
      isModalSelection: this._selection.isModal(),
      isCompactMode: false,
      myRecentItems: [],
      hideDialog: true
    };

  }
  private searchboxval: string;
  public render(): React.ReactElement<IDisplayDocumentsProps> {
    const { columns, isCompactMode, items, selectionDetails, favorites, favoriteColumns } = this.state;
    thisDuplicating = this;
    const displayStyle = {
      display: 'none',
    };
    return (
      <div className={styles.displayDocuments}>
        <div className={styles.container}>
          <div className={styles.row}>
            {/*<div className={ styles.column }>*/}
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
            <div className={'ms-Grid'}>
              <div className="search ms-Grid-row">
                <div className={'searchBocContains ' + styles.searchBocContains}>
                  <div className="ms-SearchBoxExample">
                    <div className={'ms-Grid-col ms-lg12 ' + styles.searchboxtitle}>Matters</div>
                    <div className={'ms-Grid-col ms-lg12 ' + styles.margin}>
                      <SearchBox
                        className={'ms-Grid-col ms-sm8 ms-lg9 ' + styles.searchbox}
                        placeholder="Search"
                        onSearch={newvalue => this.mattersearchresults(newvalue)}
                        onFocus={() => console.log('onFocus called')}
                        onBlur={onblurval => this.onBlurSearch(onblurval)}
                        onChange={onchangeVal => this.onChangeSearch(onchangeVal)}
                      />
                      <PrimaryButton text="Search" className={'ms-Grid-col ms-sm3 ms-lg2 ' + styles.button} onClick={() => { this.mattersearchresults(this.searchboxval); }} />
                      <div style={displayStyle} className={'ms-Grid-col ms-lg5 searchMatterResults ' + styles.searchResults}></div>
                    </div>

                  </div>
                </div>
              </div>
              <div className={'ms-Grid-row recentMatterDocs ' + styles.recent}>
                <div className={'ms-Grid-col ms-lg12 accordion3 ' + styles.accordion3}><p className={'ms-Grid-col ms-lg11 accordiontitle3 ' + styles.accordionTitle} >Recent Matters</p><div className={styles.accordionicon4}><i className={'ms-Grid-col ms-lg1 accordionicon3 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon3} aria-hidden="true"></i></div></div>
                <div className={'containsMatterRecents ' + styles.panel}>
                  <MarqueeSelection selection={this._selection}>
                    <DetailsList
                      items={items}
                      //compact={isCompactMode}
                      columns={columns}
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
              <div className={'ms-Grid-row MatterFavourites ' + styles.ClientFavourites}>
                <div className={'ms-Grid-col ms-lg12 accordion4 ' + styles.accordion4}><p className={'ms-Grid-col ms-lg11 accordiontitle4 ' + styles.accordionTitle} >Favourite Matters</p><div className={styles.accordionicon4}><i className={'ms-Grid-col ms-lg1 accordionicon4 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon4} aria-hidden="true"></i></div></div>
                <div className={'containsMatterFav ' + styles.panel}>
                  <MarqueeSelection selection={this._selection}>
                    <DetailsList
                      items={favorites}
                      //compact={isCompactMode}
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

              {/*<div className={'recentDocs '+styles.recent}>
                </div>*/}
            </div>
          </div>
        </div>
      </div>
    );
  }
  public onBlurSearch(onblurval) {
    console.log(onblurval);
    if (onblurval.path[0].value == '') {
      $('.searchMatterResults').css('display', 'none');
    }
  }
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  };
  private _deleteFav = (): void => {
    pnp.sp.site.rootWeb.lists.getByTitle('Favourite').items.getById(currFavID).delete()
      .then(res => {
        console.log('Happy');
        $(".matterfavoritebox" + currFavID).remove();
        this.setState({ hideDialog: true });
        this.getMatterFavourites('again');
      });
  };
  public onChangeSearch(onchangeText: string): void {
    this.searchboxval = onchangeText;
    if (onchangeText == "") {
      $('.searchMatterResults').html('');
      $('.searchMatterResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
      $('.searchMatterResults').removeAttr('style');
    }

  }
  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  };
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

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
  public async mattersearchresults(text: string) {
    _items = [];
    if (text != undefined && text != "") {
      $('.searchMatterResults').html('');
      let absUrl: string = this._context.pageContext.web.absoluteUrl + "/Clients/";
      var absurlIndex = absUrl.indexOf('.com');
      absUrlSubStr = absUrl.substring(0, absurlIndex + 4);
      var webAbs = this._context.pageContext.web.absoluteUrl;
      $.ajax({
        //url: absUrl + "_api/search/query?querytext='(PTLClientReference:"+text+"* OR PTLMatterReference:"+text+"* OR PTLMatterKeyDesc:"+text+"* OR PTLClientName:"+text+"*) ContentType:Managed%20Folder IsDocument:false  + path:" + absUrl + "'&selectproperties='Title,Path,RefinableString07,PTLMatterKeyDesc,PTLMatterReference,PTLClientReference,PTLClientName'&refiners=%27RefinableString07%27&rowlimit=500",
        url: absUrl + "_api/search/query?querytext='(Title:"+text+"* OR PTLMatterReference:"+text+"* OR PTLMatterKeyDesc:"+text+"*) ContentType:Managed%20Folder IsDocument:false  + path:" + absUrl + "'&selectproperties='Title,Path,RefinableString07,PTLMatterKeyDesc,PTLMatterReference,PTLClientReference,PTLClientName'&refiners=%27RefinableString07%27&rowlimit=500",
        //url: absUrl + "_api/search/query?querytext='Title:" + text + "* ContentType:Managed%20Folder (RefinableString07:a* OR RefinableString07:b* OR RefinableString07:c* OR RefinableString07:d* OR RefinableString07:e* OR RefinableString07:f* OR RefinableString07:g* OR RefinableString07:h* OR RefinableString07:i* OR RefinableString07:j* OR RefinableString07:k* OR RefinableString07:l* OR RefinableString07:m* OR RefinableString07:n* OR RefinableString07:o* OR RefinableString07:p* OR RefinableString07:q* OR RefinableString07:r* OR RefinableString07:s* OR RefinableString07:t* OR RefinableString07:u* OR RefinableString07:v* OR RefinableString07:w* OR RefinableString07:x* OR RefinableString07:y* OR RefinableString07:z* OR RefinableString07:1* OR RefinableString07:2* OR RefinableString07:3* OR RefinableString07:4* OR RefinableString07:5* OR RefinableString07:6* OR RefinableString07:7* OR RefinableString07:8* OR RefinableString07:9* OR RefinableString07:0*) IsDocument:false  + path:" + absUrl + "'&selectproperties='Title,Path,RefinableString07'",
        method: "GET",
        headers: {
          "accept": "application/json;odata=verbose",   //It defines the Data format
        },
        cache: false,
        success: function (data) {
          console.log(data);
          var strTitle = "";
          let flag: number = 0;
          if (data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.length != 0) {
            $.each(data.d.query.PrimaryQueryResult.RefinementResults.Refiners.results[0].Entries.results, function (index, value) {
              let ind: number = 0;
              //let title: String = value.Cells.results.find(x => x.Key === 'Title').Value.toLowerCase();
             /* let title: String = value.Cells.results.find(x => x.Key === 'Title').Value;
              var matterRefNo = value.Cells.results.find(x => x.Key === 'RefinableString07').Value;
              let absUrlMatterDoc: string = encodeURI(value.Cells.results.find(x => x.Key === 'OriginalPath').Value);
              ind = title.indexOf(text.toLowerCase());
              var siteassetsindx = absUrlMatterDoc.toLowerCase().indexOf('siteassets');
              var sitecollectionindx = absUrlMatterDoc.toLowerCase().indexOf('sitecollectiondocuments');*/
              let title: String = value.RefinementName;
              var titlesplit = title.split('/');
              var folderUrl = absUrl + "/" + title;
              //if (ind != -1/* && matterRefNo != '' && matterRefNo != null*/) {
              //if (siteassetsindx == -1 && sitecollectionindx == -1) {

              if (flag < searchResultsLgth) {
                strTitle += '<div class=' + styles.searchResHover + '><div><div class=' + styles.folderIconContain + '><i class="ms-Icon ms-Icon--FabricFolder" aria-hidden="true" style="font-size: 16px;padding: 2px 0 0 0;"></i></div><a class=' + styles.anchorlib + ' href=' + encodeURI(folderUrl) + '>' + titlesplit[titlesplit.length-1] + '</a></div></div>';
              }
              flag++;
              //}
              //}
            });
          }
          else {
            $('.searchMatterResults').html('');
            $('.searchMatterResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
            $('.searchMatterResults').removeAttr('style');
          }
          if (flag == 0) {
            $('.searchMatterResults').html('');
            $('.searchMatterResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
            $('.searchMatterResults').removeAttr('style');
          }
          else if (flag > searchResultsLgth) {
            var href = webAbs + seeMoreUrl + text + "*";
            strTitle += '<div class=' + styles.seemore + '><a class=' + styles.seemorelink + ' href=' + href + '>see more</a></div>';
            $('.searchMatterResults').append(strTitle);
            $('.searchMatterResults').removeAttr('style');
          }
          else {
            $('.searchMatterResults').append(strTitle);
            $('.searchMatterResults').removeAttr('style');
          }
        },
        error: function (data) {
          console.log(data);
        }
      });
    }
  }
  /*GETTING RECENTS OF MATTER FOLDERS */
  public async getMatterRecents() {
    $('.ms-SearchBox-iconContainer').css("color", "black");
    var htmlrecentstr = '';
    var absurl = this._context.pageContext.web.absoluteUrl + "/Clients/";
    _items = [];
    var title;
    var clientName, clientRef, matterKey, refinementName;
    $.ajax({
      //url: "https://xencorpdev.sharepoint.com/sites/PTL/_api/search/query?querytext=%27(Path:https://xencorpdev.sharepoint.com/sites/PTL/client)%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20ContentType:Folder%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%20NOT(FileExtension:mht%20OR%20FileExtension:aspx%20OR%20FileExtension:html%20OR%20FileExtension:htm)%27&rowlimit=50&bypassresulttypes=false&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,PTLMatterRefNo,SPWebUrl%27&sortlist=%27LastModifiedTime:descending%27&enablesorting=true&refiners='PTLMatterRefNo'",
      //url: absurl+"_api/search/query?querytext=%27(Path:"+absurl+")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20ContentType:Managed%20Folder%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%27&rowlimit=50&bypassresulttypes=false&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,RefinableString07,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,PTLMatterReference,SPWebUrl%27&sortlist=%27LastModifiedTime:descending%27&enablesorting=true&refiners='RefinableString07'",
      url: absurl + "/_api/search/query?querytext=%27ContentType:Managed%20Folder%27&properties=%27SourceName:MyRecentDocsScopeWeb,SourceLevel:SPSite%27&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,RefinableString07,PTLMatterKeyDesc,PTLMatterReference,PTLClientReference,PTLClientName,SPWebUrl%27&refiners=%27RefinableString07%27&sortlist=%27LastModifiedTime:descending%27&rowlimit=500",
      method: "GET",
      headers: {
        "accept": "application/json;odata=verbose",   //It defines the Data format
      },
      cache: false,
      success: function (recents) {
        let flag: number = 0;
        var relevantResults = recents.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;

        if (recents.d.query.PrimaryQueryResult.RefinementResults != null) {
          $.each(recents.d.query.PrimaryQueryResult.RefinementResults.Refiners.results[0].Entries.results, function (index, value) {
            //var hrefUrl = absurl + value.RefinementName;
            clientName = '';
            clientRef = '';
            matterKey = '';
            refinementName = '';
            var hrefUrl = absurl + value.RefinementName;
            refinementName = value.RefinementName;
            var titleSplit = value.RefinementName.split('/');
            var titleLgth = titleSplit.length;
            title = titleSplit[titleLgth - 1];
            //htmlrecentstr += '<div><a class=' + styles.clientAnchorlib + ' href=' + href + '>' + value.RefinementName + '</a></div>';

            if (flag < 20) {
              var queryRes = findValues(recents.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results);
              //console.log(queryRes);
              console.log(matterKey + "--" + clientName + "--" + title + "--" + clientRef);
              flag++;
              _items.push({
                name: title,
                value: title,
                iconName: '<i class="ms-Icon ms-Icon--FabricFolder" aria-hidden="true" style="font-size: 16px;padding: 2px 0 0 0;"></i>',
                type: "Folder",
                dateModified: "",
                dateModifiedValue: 0,
                fileSize: '25',
                fileSizeRaw: 30,
                displayName: "Folder",
                href: hrefUrl,
                clientName: clientName,
                clientRef: clientRef,
                matterKey: matterKey,
                matterRef: title,
              });
            }
          });
        }
        thisDuplicating.setState({ items: _items });
        //$('.containsMatterRecents').append(htmlrecentstr);
        var accordion3 = document.getElementsByClassName("accordion3");
        var i;
        for (i = 0; i < accordion3.length; i++) {
          accordion3[i].addEventListener("click", function () {
            this.classList.toggle("active");
            var panel = this.nextElementSibling;
            if (panel.style.maxHeight) {
              panel.style.maxHeight = null;
              //$('.accordionicon4').removeClass( "ms-Icon--ChevronDown" ).addClass( "ms-Icon--ChevronRight" );
              $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
              $('.accordiontitle3').css("color", "rgba(106, 191, 52, 1)");
              $('.accordionicon3').css("color", "gray");
              $('.accordionicon3').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
            } else {
              panel.style.maxHeight = panel.scrollHeight + "px";
              //$('.accordionicon4').removeClass( "ms-Icon--ChevronRight" ).addClass( "ms-Icon--ChevronDown" );
              $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
              $('.accordiontitle3, .accordionicon3').css("color", "white");
              $('.accordionicon3').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
            }
          });
        }
        $('.accordion3').trigger('click');
      },
      error: function (data) {
        console.log(data);
      }
    });
    function findValues(harddata) {
      //var matched = false;
      var data = harddata;
      var queryResult = Enumerable.from(harddata)
        .where(function (x, y) { 
          //console.log(x);
          return data[y].Cells.results;
          
        })
        //.select("$.Cells.results.value+':'+"+refinementName+"")
        .select(
          function (x, y) {
            //console.log(data[y]);
            var key = data[y].Cells.results.find(x => x.Key === 'RefinableString07').Value;
            if(key == refinementName){
              clientRef = data[y].Cells.results.find(x => x.Key === 'PTLClientReference').Value;
              matterKey = data[y].Cells.results.find(x => x.Key === 'PTLMatterKeyDesc').Value;
              clientName = data[y].Cells.results.find(x => x.Key === 'PTLClientName').Value;
              return clientRef +"@#@"+matterKey +"@#@"+clientName;
            }
          })
        .toArray();
      console.log(queryResult);
      return queryResult;
    }
  }
  /*GETTING MATTER FAVOURITES */
  public async getMatterFavourites(para) {
    var htmlStr = "";
    _favorites = [];
    pnp.sp.web.currentUser.get().then(result => {
      let currUserId: number = result.Id;
      let queryStr: string = "Author eq " + currUserId + "and PTLCategory eq 'Matter'";
      pnp.sp.web.lists.getByTitle("Favourite")
        .items
        .select(
          "Id",
          "Title",
          "PTLItemURL",
          "PTLCategory",
          "PTLFavouriteMetadata"
        )
        .filter(queryStr)
        .get()
        .then((data) => {
          //console.log('happy');
          $.each(data, function (index, value) {
            //htmlStr += '<div><a class='+styles.clientAnchorlib+' href='+value.PTLItemURL+'>'+value.Title+'</a></div>';
            //htmlStr += '<div id="matterfavoritebox' + index + '" class="ms-Grid-col ms-lg5 matterfavoritebox' + value.ID + ' ' + styles.favoritebox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg6 clientAnchorlib"><div class="clientAnchorlib ' + styles.clientAnchorlib + '">' + value.Title + '</div></div><div title="Launch Favorite" class="drillbox ms-Grid-col ms-lg2 drillbox ' + styles.launchbox + '"><a href="' + value.PTLItemURL + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></a></div><div title="Delete Favorite" class="deletebox ms-Grid-col ms-lg2 ' + styles.deletebox + '" value = "' + value.ID + '"><i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></div></div>';
            //htmlStr += '<div id="favoritebox' + index + '" class="ms-Grid-col ms-lg5 favoritebox'+value.ID+' ' + styles.favoritebox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg6 clientAnchorlib"><div class="clientAnchorlib ' + styles.clientAnchorlib + '">' + value.Title + '</div></div><div title="Launch Favorite" class="drillbox ms-Grid-col ms-lg2 drillbox ' + styles.launchbox + '"><a href="'+value.PTLItemURL+'"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></a></div><div title="Delete Favorite" class="deletebox ms-Grid-col ms-lg2 '+styles.deletebox+'" value = "'+value.ID+'"><i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></div></div>';
            var metadata = [];
            try{
              metadata = value.PTLFavouriteMetadata.split('#@#');
            }catch(e){
              //Nothing
            }
            
            
            _favorites.push({
              name: value.Title,
              value: value.Title,
              iconName: '<i class="ms-Icon ms-Icon--FabricFolder" aria-hidden="true" style="font-size: 16px;padding: 2px 0 0 0;"></i>',
              type: '<a class="link" data-value="' + value.PTLItemURL + '"><div title="Launch Favourite" class="drillbox ms-Grid-col ms-lg2 drillbox ' + styles.launchbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></a><div title="Delete Favourite" class="deletebox ms-Grid-col ms-lg2 ' + styles.deletebox + '" value = "' + value.ID + '"><i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></div>',
              href: value.PTLItemURL,
              clientName: metadata[0],
              clientRef: metadata[1],
              matterKey: metadata[2],
              matterRef: metadata[3],
            });
          });
          thisDuplicating.setState({ favorites: _favorites });
          //$('.containsMatterFav').append(htmlStr);
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
            var accordion4 = document.getElementsByClassName("accordion4");
            var i;
            for (i = 0; i < accordion4.length; i++) {
              accordion4[i].addEventListener("click", function () {
                this.classList.toggle("active");
                var panel = this.nextElementSibling;
                if (panel.style.maxHeight) {
                  panel.style.maxHeight = null;
                  //$('.accordionicon4').removeClass( "ms-Icon--ChevronDown" ).addClass( "ms-Icon--ChevronRight" );
                  $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
                  $('.accordiontitle4').css("color", "rgba(106, 191, 52, 1)");
                  $('.accordionicon4').css("color", "gray");
                  $('.accordionicon4').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
                } else {
                  panel.style.maxHeight = panel.scrollHeight + "px";
                  //$('.accordionicon4').removeClass( "ms-Icon--ChevronRight" ).addClass( "ms-Icon--ChevronDown" );
                  $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
                  $('.accordiontitle4, .accordionicon4').css("color", "white");
                  $('.accordionicon4').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
                }
              });
            }
          }

        });
    });
  }

}

