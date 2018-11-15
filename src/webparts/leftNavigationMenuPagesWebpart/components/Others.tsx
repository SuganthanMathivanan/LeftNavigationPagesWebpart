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
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    SelectionMode,
    IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import Dropdown from 'react-dropdown';
import 'react-dropdown/style.css'
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import {
    css,
    DocumentCard,
    DocumentCardLocation,
    DocumentCardPreview,
    DocumentCardTitle,
    DocumentCardActivity,
    Spinner,
    SpinnerSize,
} from 'office-ui-fabric-react';
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

}
let _items: IDocument[] = [];
var _dropdownItems = [];
let _favorites: IDocument[] = [];
export interface IDetailsListDocumentsExampleState {
    columns: IColumn[];
    //favoriteColumns: IColumn[];
    items: IDocument[];
    dropdownItems;
    favorites: IDocument[];
    selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    myRecentItems: IMyRecentItems[];
    hideDialog: boolean;
    selectedKey: string;
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
    fileSizeRaw: number;*/
    displayName: string;
}
var dropdownflag = 0;
var valOfDropdown;
var thisDuplicating;
let searchResultsLgth: number = 2;
var absUrlSubStr;
var defaultSelectedKey;
let seeMoreUrl: string = "/Search/Pages/results.aspx?k=ContentType:Managed%20Folder%20Title:";
export default class OtherDocuments extends React.Component<IDisplayDocumentsProps, IDetailsListDocumentsExampleState> {
    private _context: WebPartContext;
    private _currentWebUrl: string;
    private _selection: Selection;
    //public searchResultsLgth:number;
    constructor(props: IDisplayDocumentsProps) {
        super(props);

        this._context = props.context;
        this.getlibraries();
        this._currentWebUrl = this._context.pageContext.web.absoluteUrl;
        //this.searchResultsLgth = 2;
        /*this.state = {
            myRecentItems: []
        };*/


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
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IDocument) => {
                    return <a href={item.href} className={styles.titleAnchorStyle}>{item.name}</a>;
                }
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
            selectedKey: undefined,
            items: [],
            dropdownItems: [],
            favorites: [],
            columns: _columns,
            //favoriteColumns: _favoriteColumns,
            selectionDetails: this._getSelectionDetails(),
            isModalSelection: this._selection.isModal(),
            isCompactMode: false,
            myRecentItems: [],
            hideDialog: true
        };

    }
    private searchboxval: string;
    public render(): React.ReactElement<IDisplayDocumentsProps> {
        const { columns, isCompactMode, items, selectionDetails, dropdownItems } = this.state;
        //const{ selectedKey } =  this.props;
        const displayStyle = {
            display: 'none',
        };

        thisDuplicating = this;
        /*const options = [
            { key: "", text: "Select..." },
            { key: 'Client', text: "Client" },
            { key: 'Matter', text: "Matter" }
        ];*/
        /*const options = [
            { value: 'one', label: 'One' },
            { value: 'two', label: 'Two', className: 'myOptionClassName' }
          ]*/
        const options = _dropdownItems;
        const defaultOption = options[0]

        return (
            <div className={styles.displayDocuments}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        {/*<div className={'ms-Grid '+ styles.column }>-->*/}
                        <div className={'ms-Grid'} dir="ltr">
                            <div className="search ms-Grid-row">
                                <div className={'searchBocContains ' + styles.searchBocContains}>
                                    <div className="ms-SearchBoxExample">
                                        <div className={'ms-Grid-col ms-lg12 ' + styles.searchboxtitle}>Others</div>
                                        <div className={'ms-Grid-col ms-lg12'}>
                                            <Dropdown
                                                className="libDropdown"
                                                //efaultSelectedKey="Entity1"
                                                value={defaultOption}
                                                placeholder="Select an option"
                                                options={options}
                                            />
                                        </div>
                                        <div className={"blockUI blockOverlay " + styles.blockOverlay}></div>
                                        <div className={"blockUI blockMsg blockPage " + styles.blockOverlay}>
                                            <Spinner size={SpinnerSize.large} label="loading..." />
                                        </div>
                                        <div className={'ms-Grid-col ms-lg12 ' + styles.margin}>
                                            <SearchBox
                                                className={'ms-Grid-col ms-sm8 ms-lg9 ' + styles.searchbox}
                                                placeholder="Search"
                                                onSearch={newvalue => this.othersearchresults(newvalue)}
                                                //onFocus={() => console.log('onFocus called')}
                                                onBlur={onblurval => this.onBlurSearch(onblurval)}
                                                onChange={onchangeVal => this.onChangeSearch(onchangeVal)}
                                            />
                                            <PrimaryButton className={'ms-Grid-col ms-sm3 ms-lg2 ' + styles.button} text="Search" onClick={() => { this.othersearchresults(this.searchboxval); }} />
                                            <div style={displayStyle} className={'ms-Grid-col ms-lg5 opposearchResults ' + styles.searchResults}>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className={'ms-Grid-row recentDocs ' + styles.recent}>
                                <div className={'ms-Grid-col ms-lg12 accordion1 ' + styles.accordion1}><p className={'ms-Grid-col ms-lg11 accordiontitle1 ' + styles.accordionTitle} >Other Libraries</p><div className={styles.accordionicon4}><i className={'ms-Grid-col ms-lg1 accordionicon1 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon1} aria-hidden="true"></i></div></div>
                                <div className={'containsoppoRecents ' + styles.panel}>
                                    <MarqueeSelection selection={this._selection}>
                                        <DetailsList
                                            items={items}
                                            compact={isCompactMode}
                                            columns={columns}
                                            selectionMode={this.state.isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                                            setKey="set"
                                            layoutMode={DetailsListLayoutMode.justified}
                                            isHeaderVisible={true}
                                            //selection={this._selection}
                                            selectionPreservedOnEmptyClick={true}
                                            //onItemInvoked={this._onItemInvoked}
                                            enterModalSelectionOnTouch={true}
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

    private onChange(ev, item): void {
        this.setState({
            //items: this.state.items.concat([item]),
            selectedKey: item.text
        });
        //this._OnchangedCategory = item.key;
    };
    public onBlurSearch(onblurval) {
        console.log(onblurval);
        if (onblurval.path[0].value == '') {
            $('.opposearchResults').css('display', 'none');
        }
    }

    public onChangeSearch(onchangeText: string): void {
        this.searchboxval = onchangeText;
        if (onchangeText == "") {
            $('.opposearchResults').html('');
            $('.opposearchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
            $('.opposearchResults').removeAttr('style');
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
    /*public async dropdown() {
        var p = document.getElementById("Dropdown23-option");
                    var i;
        //for (i = 0; i < accordion1.length; i++) {
        if (p != null) {
                        p[i].addEventListener("change", function () {
                            console.log('happy');

                        });
                    }
        if (p != null) {
                        p[i].addEventListener("DOMSubtreeModified", function () {
                            alert('changed')
                        });
                    }
                    //}
                }*/
    public async othersearchresults(text: string) {
        if (text != undefined && text != "") {
            $('.opposearchResults').html('');
            var absUrlOther = this._context.pageContext.web.absoluteUrl + "/Others";
            console.log(absUrlOther);
            var absurlIndex = absUrlOther.indexOf('.com');
            absUrlSubStr = absUrlOther.substring(0, absurlIndex + 4);
            var webAbs = this._context.pageContext.web.absoluteUrl;
            var libUrl = absUrlOther + '/' + $('.Dropdown-placeholder.is-selected').text();
            var icon = absUrlSubStr + "/_layouts/15/images/itdl.png?rev=44";
            $.ajax({
                //url: absUrl + "_api/search/query?querytext='(PTLClientReference:"+text+"* OR PTLMatterReference:"+text+"* OR PTLMatterKeyDesc:"+text+"* OR PTLClientName:"+text+"*) ContentType:Managed%20Folder IsDocument:false  + path:" + absUrl + "'&selectproperties='Title,Path,RefinableString07,PTLMatterKeyDesc,PTLMatterReference,PTLClientReference,PTLClientName'&refiners=%27RefinableString07%27&rowlimit=500",
                //url: absUrl + "/_api/search/query?querytext='(Title:" + text + "*) ContentType:Managed%20Folder IsDocument:false  + path:" + libUrl + "'&selectproperties='Title,Path,RefinableString07,PTLMatterKeyDesc,PTLMatterReference,PTLClientReference,PTLClientName'&refiners=%27RefinableString07%27&rowlimit=500",
                url: absUrlOther + "/_api/search/query?querytext='(Title:"+text+"* OR PTLFriendlyName:" + text + "*)  path:" + libUrl + "'&selectproperties='Title,Path,RefinableString07,PTLMatterKeyDesc,PTLMatterReference,PTLClientReference,PTLClientName'&refiners=%27RefinableString07%27&rowlimit=500",
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
                            var folderUrl = absUrlOther + "/" + title;
                            //if (ind != -1/* && matterRefNo != '' && matterRefNo != null*/) {
                            //if (siteassetsindx == -1 && sitecollectionindx == -1) {

                            if (flag < searchResultsLgth) {
                                strTitle += '<div class=' + styles.searchResHover + '><div><div class=' + styles.folderIconContain + '><i class="ms-Icon ms-Icon--FabricFolder" aria-hidden="true" style="font-size: 16px;padding: 2px 0 0 0;"></i></div><a class=' + styles.anchorlib + ' href=' + encodeURI(folderUrl) + '>' + titlesplit[titlesplit.length - 1] + '</a></div></div>';
                            }
                            flag++;
                            //}
                            //}
                        });
                    }
                    else {
                        $('.opposearchResults').html('');
                        $('.opposearchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
                        $('.opposearchResults').removeAttr('style');
                    }
                    if (flag == 0) {
                        $('.opposearchResults').html('');
                        $('.opposearchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
                        $('.opposearchResults').removeAttr('style');
                    }
                    else if (flag > searchResultsLgth) {
                        var href = webAbs + seeMoreUrl + text + "*";
                        strTitle += '<div class=' + styles.seemore + '><a class=' + styles.seemorelink + ' href=' + href + '>see more</a></div>';
                        $('.opposearchResults').append(strTitle);
                        $('.opposearchResults').removeAttr('style');
                    }
                    else {
                        $('.opposearchResults').append(strTitle);
                        $('.opposearchResults').removeAttr('style');
                    }
                },
                error: function (data) {
                    console.log(data);
                }
            });
            /*pnp.sp.site.getDocumentLibraries(absUrl)
                .then((data) => {
                    var strTitle = "";
                    let flag: number = 0;
                    var count
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
                        $('.opposearchResults').html('');
                        $('.opposearchResults').append('<div class=' + styles.nosearchres + '>No search Results</div>');
                        $('.opposearchResults').removeAttr('style');
                    }
                    else if (flag > searchResultsLgth) {
                        var href = webAbs + seeMoreUrl + text + "*";
                        strTitle += '<div class=' + styles.seemore + '><a class=' + styles.seemorelink + ' href=' + href + '>see more</a></div>';
                        $('.opposearchResults').append(strTitle);
                        $('.opposearchResults').removeAttr('style');
                    }
                    else {
                        $('.opposearchResults').append(strTitle);
                        $('.opposearchResults').removeAttr('style');
                    }
                }).catch(function (err) {
                    alert(err);
                });*/
        }
    }
    public async getlibraries() {
        $('.ms-SearchBox-iconContainer').css("color", "black")
        var htmlrecentstr = '';
        let absUrl: string = this._context.pageContext.web.absoluteUrl + "/Others";
        var absurlIndex = absUrl.indexOf('.com');
        absUrlSubStr = absUrl.substring(0, absurlIndex + 4);
        _items = [];
        _dropdownItems = [];
        pnp.sp.site.getDocumentLibraries(absUrl)
            .then((data) => {
                $.each(data, function (index, value) {
                    var hrefUrl = value.AbsoluteUrl;
                    //htmlrecentstr += '<a  href="' + href + '"><div id="clientrecentbox' + index + '" class="ms-Grid-col ms-lg5 clientrecentbox ' + styles.clientrecentbox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg7 clientAnchorlib"><div class="clientAnchorlib ' + styles.clientAnchorlib + '">' + value.RefinementName + '</div></div><div class="drillbox ms-Grid-col ms-lg3 drillbox ' + styles.drillbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></div></a>';
                    /* htmlStr += '<div id="favoritebox' + index + '" class="ms-Grid-col ms-lg5 favoritebox ' + styles.favoritebox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div>
                     <div class="ms-Grid-col ms-lg7 clientAnchorlib"><a class="clientAnchorlib ' + styles.clientAnchorlib + '" href="' + value.PTLItemURL + '">' + value.Title + '</a></div><div class="drillbox ms-Grid-col ms-lg3 drillbox ' + styles.drillbox + '">
                     <i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></div>';*/
                    _items.push({
                        name: value.Title,
                        value: value.Title,
                        iconName: absUrlSubStr + "/_layouts/15/images/itdl.png?rev=44",
                        type: "Document Library",
                        href: hrefUrl,
                        displayName: "Document Library",
                    });
                    _dropdownItems.push(
                        //{ key: value.Title, text: value.Title },
                        { value: value.Title, label: value.Title },
                    );
                });
                //$('.containsoppoRecents').append(htmlrecentstr);
                //thisDuplicating.setState({ dropdownItems: _dropdownItems });

                thisDuplicating.setState({ items: _items, dropdownItems: _dropdownItems });
                var span = document.createElement('span');
                span.textContent = "Hello!";
                var s = document.getElementsByClassName("ms-Dropdown-title");
                //this.render();


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
                            panel.style.display = "block";
                            $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
                            $('.accordiontitle1, .accordionicon1').css("color", "white");
                            $('.accordionicon1').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
                        }
                    });
                }
                $('.accordion1').trigger('click');
                /*var dropDown = document.getElementsByClassName("ms-Dropdown-title");
                var i;
                for (i = 0; i < dropDown.length; i++) {
                    dropDown[i].addEventListener("DOMSubtreeModified", function (e) {
                        valOfDropdown = $('#Dropdown23-option>span').text();
                        libUrl = absUrl + "/" + valOfDropdown;

                    });
                }*/
            }).catch(function (err) {
                alert(err);
            });

    }

}

function test() {

    console.log('happy');
}