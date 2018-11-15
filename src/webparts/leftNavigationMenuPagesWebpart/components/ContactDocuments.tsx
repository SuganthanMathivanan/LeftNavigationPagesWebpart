//import jQuery from 'jQuery';
import * as React from 'react';
import styles from './LeftNavigationMenuPagesWebpart.module.scss';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import {
    List,
    PrimaryButton
} from 'office-ui-fabric-react';
//import { IDisplayDocumentsProps } from './IDisplayDocumentsProps';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import {
    css,
    DocumentCard,
    DocumentCardLocation,
    DocumentCardPreview,
    DocumentCardTitle,
    DocumentCardActivity,
    Spinner
} from 'office-ui-fabric-react';
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    SelectionMode,
    IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import * as Enumerable from 'linq';
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

/*const fileIcons: { name: string }[] = [
  { name: 'accdb' },
  { name: 'csv' },
  { name: 'docx' },
  { name: 'dotx' },
  { name: 'mpp' },
  { name: 'mpt' },
  { name: 'odp' },
  { name: 'ods' },
  { name: 'odt' },
  { name: 'one' },
  { name: 'onepkg' },
  { name: 'onetoc' },
  { name: 'potx' },
  { name: 'ppsx' },
  { name: 'pptx' },
  { name: 'pub' },
  { name: 'vsdx' },
  { name: 'vssx' },
  { name: 'vstx' },
  { name: 'xls' },
  { name: 'xlsx' },
  { name: 'xltx' },
  { name: 'xsn' }
];*/

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
    friendlyName: any;
}
let searchResultsLgth: number = 2;
var absUrlSubStr;
let seeMoreUrl: string = "/Search/Pages/results.aspx?k=ContentType:Managed%20Folder%20IsDocument:false%20Title:";
var currFavID = 0;
var thisDuplicating;
export default class ContactDocuments extends React.Component<IDisplayDocumentsProps, IDetailsListDocumentsExampleState> {
    private _context: WebPartContext;
    private _currentWebUrl: string;
    private _selection: Selection;
    constructor(props: IDisplayDocumentsProps) {
        super(props);
        this._context = props.context;
        this._currentWebUrl = this._context.pageContext.web.absoluteUrl;
        /*this.state = {
            myRecentItems: [],
            hideDialog: true
        };*/
        this.getContactRecents();
        this.getContactFavourites('new');
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

                },
            },
            {
                key: 'column3',
                name: 'Friendly Name',
                fieldName: 'friendlyname',
                minWidth: 130,
                maxWidth: 130,
                isResizable: true,
                isCollapsable: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span className={styles.authortext}>{item.friendlyName}</span>;

                },
                isPadded: true
            }/*,
            {
              key: 'column3',
              name: 'Type',
              fieldName: 'type',
              minWidth: 130,
              maxWidth: 130,
              isResizable: true,
              isCollapsable: true,
              data: 'string',
              onColumnClick: this._onColumnClick,
              onRender: (item: IDocument) => {
                return <span className={styles.authortext}>{item.type}</span>;
      
              },
              isPadded: true
            }*/
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
                minWidth: 350,
                maxWidth: 350,
                isRowHeader: true,
                //isResizable: true,
                //isSorted: true,
                //isSortedDescending: false,
                //sortAscendingAriaLabel: 'Sorted A to Z',
                //sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'column3',
                name: 'Friendly Name',
                fieldName: 'friendlyname',
                minWidth: 130,
                maxWidth: 130,
                isResizable: true,
                isCollapsable: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span className={styles.authortext}>{item.friendlyName}</span>;

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
                                        <div className={'ms-Grid-col ms-lg12 ' + styles.searchboxtitle}>Contacts</div>
                                        <div className={'ms-Grid-col ms-lg12 ' + styles.margin}>
                                            <SearchBox
                                                className={'ms-Grid-col ms-sm8 ms-lg9 ' + styles.searchbox}
                                                placeholder="Search"
                                                onSearch={newvalue => this.contactsearchresults(newvalue)}
                                                onFocus={() => console.log('onFocus called')}
                                                onBlur={onblurval => this.onBlurSearch(onblurval)}
                                                onChange={onchangeVal => this.onChangeSearch(onchangeVal)}
                                            />
                                            <PrimaryButton className={'ms-Grid-col ms-sm3 ms-lg2 ' + styles.button} text="Search" onClick={() => { this.contactsearchresults(this.searchboxval); }} />
                                            <div style={displayStyle} className={'ms-Grid-col ms-lg5 ContactSearchResults ' + styles.searchResults}>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className={'ms-Grid-row recentDocs ' + styles.recent}>
                                <div className={'ms-Grid-col ms-lg12 accordion1 ' + styles.accordion1}><p className={'ms-Grid-col ms-lg11 accordiontitle1 ' + styles.accordionTitle} >Recent Contacts</p><div className={styles.accordionicon4}><i className={'ms-Grid-col ms-lg1 accordionicon1 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon1} aria-hidden="true"></i></div></div>
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
                                        //selection={this._selection}
                                        //selectionPreservedOnEmptyClick={true}
                                        //onItemInvoked={this._onItemInvoked}
                                        //enterModalSelectionOnTouch={true}
                                        />
                                    </MarqueeSelection>
                                </div>
                            </div>
                            <div className={'ms-Grid-row ClientFavourites ' + styles.ClientFavourites}>
                                <div className={'ms-Grid-col ms-lg12 accordion2 ' + styles.accordion2}><p className={'ms-Grid-col ms-lg11 accordiontitle2 ' + styles.accordionTitle} >Favourite Contacts</p><div className={styles.accordionicon4}><i className={'ms-Grid-col ms-lg1 accordionicon2 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon2} aria-hidden="true"></i></div></div>
                                <div className={'containsFav ' + styles.panel}>
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
                        </div>
                    </div>
                </div>
            </div>
        );
    }
    public onBlurSearch(onblurval) {
        console.log(onblurval);
        if (onblurval.path[0].value == '') {
            $('.ContactSearchResults').css('display', 'none');
        }
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
                this.getContactFavourites('again');
            });
    };
    private _showDialog = (): void => {
        this.setState({ hideDialog: false });
    };
    public onChangeSearch(onchangeText: string): void {
        this.searchboxval = onchangeText;
        if (onchangeText == "") {
            $('.ContactSearchResults').html('');
            $('.ContactSearchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
            $('.ContactSearchResults').removeAttr('style');
        }

    }
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
    public async contactsearchresults(text: string) {
        if (text != undefined && text != "") {
            $('.ContactSearchResults').html('');
            let absUrl: string = this._context.pageContext.web.absoluteUrl + '/Contacts';
            var absurlIndex = absUrl.indexOf('.com');
            absUrlSubStr = absUrl.substring(0, absurlIndex + 4);
            var webAbs = this._context.pageContext.web.absoluteUrl;
            $.ajax({
                url: absUrl + "/_api/search/query?querytext='(Title:"+text+"* OR PTLFriendlyName:" + text + "*)'&selectproperties='Title,Path,RefinableString07,PTLFriendlyName,PTLClientReference'&refiners=%27RefinableString07%27&rowlimit=500",
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
                            let title: String = value.RefinementName;
                            var titlesplit = title.split('/');
                            var folderUrl = absUrl + "/" + title;
                            /*let title: String = value.Cells.results.find(x => x.Key === 'Title').Value.toLowerCase();
                            var matterRefNo = value.Cells.results.find(x => x.Key === 'RefinableString07').Value;
                            let absUrlMatterDoc: string = encodeURI(value.Cells.results.find(x => x.Key === 'OriginalPath').Value);
                            ind = title.indexOf(text.toLowerCase());
                            var siteassetsindx = absUrlMatterDoc.toLowerCase().indexOf('siteassets');
                            var sitecollectionindx = absUrlMatterDoc.toLowerCase().indexOf('sitecollectiondocuments');*/
                            //if (ind != -1 /*&& matterRefNo != '' && matterRefNo != null*/) {
                            //if (siteassetsindx == -1 && sitecollectionindx == -1) {

                            if (flag < searchResultsLgth) {
                                strTitle += '<div class=' + styles.searchResHover + '><div><div class=' + styles.folderIconContain + '><i class="ms-Icon ms-Icon--FabricFolder" aria-hidden="true" style="font-size: 16px;padding: 2px 0 0 0;"></i></div><a class=' + styles.anchorlib + ' href=' + folderUrl + '>' + titlesplit[titlesplit.length-1] + '</a></div></div>';
                            }
                            flag++;
                            //}
                            //}
                        });
                    }
                    else {
                        $('.ContactSearchResults').html('');
                        $('.ContactSearchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
                        $('.ContactSearchResults').removeAttr('style');
                    }
                    if (flag == 0) {
                        $('.ContactSearchResults').html('');
                        $('.ContactSearchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
                        $('.ContactSearchResults').removeAttr('style');
                    }
                    else if (flag > searchResultsLgth) {
                        var href = webAbs + seeMoreUrl + text + "*";
                        strTitle += '<div class=' + styles.seemore + '><a class=' + styles.seemorelink + ' href=' + href + '>see more</a></div>';
                        $('.ContactSearchResults').append(strTitle);
                        $('.ContactSearchResults').removeAttr('style');
                    }
                    else {
                        $('.ContactSearchResults').append(strTitle);
                        $('.ContactSearchResults').removeAttr('style');
                    }
                },
                error: function (data) {
                    console.log(data);
                }
            });
        }
    }
    public async getContactRecents() {
        $('.ms-SearchBox-iconContainer').css("color", "black")
        var htmlrecentstr = '';
        let absUrl: string = this._context.pageContext.web.absoluteUrl;
        var serverRelUrl = this._context.pageContext.web.serverRelativeUrl;
        var refinementName, PTLFriendlyName;
        _items = [];
        $.ajax({
            //url: this._context.pageContext.web.absoluteUrl+"/_api/search/query?querytext=%27(Path:"+this._context.pageContext.web.absoluteUrl+")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20ContentType:Managed%20Folder%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%20NOT(FileExtension:mht%20OR%20FileExtension:aspx%20OR%20FileExtension:html%20OR%20FileExtension:htm)%27&rowlimit=50&bypassresulttypes=false&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,SPWebUrl,RefinableString07%27&sortlist=%27LastModifiedTime:descending%27&enablesorting=true&refiners='RefinableString07'",
            url: absUrl + "/Contacts/_api/search/query?querytext=%27*%27&properties=%27SourceName:MyRecentDocsScopeWeb,SourceLevel:SPSite%27&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,LastModifiedTime,ModifiedBy,LinkingUrl,SiteTitle,ParentLink,DocumentPreviewMetadata,ListID,ListItemID,SPSiteURL,SiteID,WebId,UniqueID,RefinableString07,PTLFriendlyName,SPWebUrl%27&refiners=%27RefinableString07%27&sortlist=%27LastModifiedTime:descending%27",
            method: "GET",
            headers: {
                "accept": "application/json;odata=verbose",   //It defines the Data format
            },
            cache: false,
            success: function (recents) {
                let flag: number = 0;
                if (recents.d.query.PrimaryQueryResult.RefinementResults != null) {
                    $.each(recents.d.query.PrimaryQueryResult.RefinementResults.Refiners.results[0].Entries.results, function (index, value) {
                        //var href = "https://xencorpdev.sharepoint.com/sites/PTL/client/" + value.RefinementName;
                        var hrefUrl = absUrl + "/Contacts/" + value.RefinementName + "?id=" + serverRelUrl + "/Contacts/" + value.RefinementName;
                        var titleSplit = value.RefinementName.split('/');
                        var titleLgth = titleSplit.length;
                        var title = titleSplit[titleLgth - 1];
                        refinementName = value.RefinementName;
                        //working htmlrecentstr += '<a  href="' + href + '"><div id="clientrecentbox' + index + '" class="ms-Grid-col ms-lg5 clientrecentbox ' + styles.clientrecentbox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg7 clientAnchorlib"><div class="clientAnchorlib ' + styles.clientAnchorlib + '">' + value.RefinementName + '</div></div><div class="drillbox ms-Grid-col ms-lg3 drillbox ' + styles.drillbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></div></a>';
                        /* htmlStr += '<div id="favoritebox' + index + '" class="ms-Grid-col ms-lg5 favoritebox ' + styles.favoritebox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div>
                         <div class="ms-Grid-col ms-lg7 clientAnchorlib"><a class="clientAnchorlib ' + styles.clientAnchorlib + '" href="' + value.PTLItemURL + '">' + value.Title + '</a></div><div class="drillbox ms-Grid-col ms-lg3 drillbox ' + styles.drillbox + '">
                         <i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></div>';*/

                        if (flag < 20) {
                            var queryRes = findValues(recents.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results);
                            flag++;
                            _items.push({
                                name: title,
                                value: title,
                                iconName: '<i class="ms-Icon ms-Icon--FabricFolder" aria-hidden="true" style="font-size: 16px;padding: 2px 0 0 0;"></i>',
                                type: "Folder",
                                href: hrefUrl,
                                friendlyName: PTLFriendlyName
                            });
                        }
                    });
                }
                //$('.containsRecents').append(htmlrecentstr);
                thisDuplicating.setState({ items: _items });
                function overhandler(ev) {
                    var target = $(ev.currentTarget);
                    var elId = target.attr('id');
                    var childparty = ev.currentTarget.childNodes[0];
                    if (target.is(".clientrecentbox")) {
                        $('#' + elId).css("background-color", "#6abf34");
                        $(ev.currentTarget.childNodes[0]).css("color", "white");
                        $(ev.currentTarget.children[1].childNodes[0]).css("color", "white");
                        $(ev.currentTarget.childNodes[2]).css("color", "white");
                    }
                }
                $(".clientrecentbox").mouseover(overhandler);
                $(".clientrecentbox").mouseleave(leavehandler);
                function leavehandler(ev) {
                    var target = $(ev.currentTarget);
                    var elId = target.attr('id');
                    if (target.is(".clientrecentbox")) {
                        $('#' + elId).css("background-color", "#f8f8f8");
                        $(ev.currentTarget.childNodes[0]).css("color", "#373737");
                        $(ev.currentTarget.children[1].childNodes[0]).css("color", "#373737");
                        $(ev.currentTarget.childNodes[2]).css("color", "#6abf34");
                    }
                }

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
        function findValues(harddata) {
            //var matched = false;
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
                        var key = data[y].Cells.results.find(x => x.Key === 'RefinableString07').Value;
                        if (key == refinementName) {
                            PTLFriendlyName = data[y].Cells.results.find(x => x.Key === 'PTLFriendlyName').Value;
                            return PTLFriendlyName;
                        }
                    })
                .toArray();
            console.log(queryResult);
            return queryResult;
        }
    }
    public async getContactFavourites(para) {
        var htmlStr = "";
        _favorites = [];
        pnp.sp.web.currentUser.get().then(result => {
            let currUserId: number = result.Id;
            let queryStr: string = "Author eq " + currUserId + "and PTLCategory eq 'Contacts'";
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
                    $.each(data, function (index, value) {
                        htmlStr += '<div id="favoritebox' + index + '" class="ms-Grid-col ms-lg5 favoritebox' + value.ID + ' ' + styles.favoritebox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg6 clientAnchorlib"><div class="clientAnchorlib ' + styles.clientAnchorlib + '">' + value.Title + '</div></div><div title="Launch Favorite" class="drillbox ms-Grid-col ms-lg2 drillbox ' + styles.launchbox + '"><a href="' + value.PTLItemURL + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></a></div><div title="Delete Favorite" class="deletebox ms-Grid-col ms-lg2 ' + styles.deletebox + '" value = "' + value.ID + '"><i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></div></div>';
                        _favorites.push({
                            name: value.Title,
                            value: value.Title,
                            iconName: '<i class="ms-Icon ms-Icon--FabricFolder" aria-hidden="true" style="font-size: 16px;padding: 2px 0 0 0;"></i>',
                            type: '<a class="link" data-value="' + value.PTLItemURL + '"><div title="Launch Favourite" class="drillbox ms-Grid-col ms-lg2 drillbox ' + styles.launchbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></a><div title="Delete Favourite" class="deletebox ms-Grid-col ms-lg2 ' + styles.deletebox + '" value = "' + value.ID + '"><i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i></div>',
                            href: "",
                            friendlyName: value.PTLFavouriteMetadata,
                        });
                    });
                    thisDuplicating.setState({ favorites: _favorites });
                    //$('.containsFav').append(htmlStr);
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
                            });
                        }
                    }

                });
        });
    }
}

