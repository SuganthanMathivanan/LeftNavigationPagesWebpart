
import * as React from 'react';
import styles from './LeftNavigationMenuPagesWebpart.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

//import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    SelectionMode,
    IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
//  var fa = require('fa');
//import './DetailsListExample.scss';
//import { lorem } from 'office-ui-fabric-react/lib/utilities/exampleData';
//import { GraphHttpClient, HttpClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';
import * as pnp from 'sp-pnp-js';
import Web from 'sp-pnp-js';
import * as $ from 'jquery';

export interface IMyDocumentsProps {
    context: WebPartContext;
    description: string;
}
let _items: IDocument[] = [];
let _checkedout: IDocument[] = [];
let _followed: IDocument[] = [];
let _recentEmail: IDocument[] = [];


export interface IDetailsListDocumentsExampleState {
    columns: IColumn[];
    checkOutColumns: IColumn[];
    followColumns: IColumn[]
    items: IDocument[];
    followed: IDocument[];
    checkedout: IDocument[];
    recentEmail: IDocument[];
    selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
}

export interface IDocument {
    [key: string]: any;
    name: String;
    value: String;
    folderPath: string;
    iconName: string;
    author: String;
    dateModified: String;
    displayName: string;
    href: any;
    clientReference: any;
    clientName: any;
    matterReference: any;
    matterKeyDesc: any;
    recordName: any;
}
export interface ICheckedOutDocument {
    [key: string]: any;
    name: String;
    value: String;
    iconName: string;
    href: any;
}
var thisDuplicating;
var rowLimit = 500;
var nameOfTenant;
var followedLink = '';

export default class MyDocuments extends React.Component<IMyDocumentsProps, IDetailsListDocumentsExampleState> {
    private _context: WebPartContext;
    private _selection: Selection;
    constructor(props: IMyDocumentsProps, context) {
        super(props);
        this._context = props.context;
        this.getMyRecentDocs();
        this.getCheckedOutDocuments();
        this.getFollowedDocuments();
        this.getrecentemails();
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
                minWidth: 20,
                maxWidth: 20,
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    //return <img src={item.iconName} className={'DetailsListExample-documentIconImage'} />;
                    return <span dangerouslySetInnerHTML={{ __html: item.iconName }}></span>;
                }
            }, {
                key: 'column11',
                name: ' ',
                fieldName: 'folderpath',
                minWidth: 16,
                maxWidth: 16,
                isRowHeader: true,
                //isResizable: true,
                //isSorted: true,
                //isSortedDescending: true,
                //sortAscendingAriaLabel: 'Sorted A to Z',
                //sortDescendingAriaLabel: 'Sorted Z to A',
                //onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IDocument) => {
                    return <a href={item.folderPath} className={styles.titleAnchorStyle}><div><i title="Launch Folder" className="ms-Icon ms-Icon--OpenInNewWindow" aria-hidden="true"></i></div></a>;
                }
            },
            {
                key: 'column2',
                name: 'Name',
                fieldName: 'name',
                minWidth: 180,
                maxWidth: 180,
                isRowHeader: true,
                //isResizable: true,
                //isSorted: true,
                //isSortedDescending: true,
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
                name: 'Date Modified',
                fieldName: 'dateModifiedValue',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: false,
                //onColumnClick: this._onColumnClick,
                data: 'number',
                onRender: (item: IDocument) => {
                    return <span>{item.dateModified}</span>;
                },
                isPadded: true
            },
            {
                key: 'column4',
                name: 'Author',
                fieldName: 'author',
                minWidth: 20,
                maxWidth: 20,
                //isResizable: false,
                //isCollapsable: false,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <div className="circleContains"><div className={styles.authorcircle}><div className={styles.authorFirstLetter}>{item.author}</div></div></div>;

                },
                isPadded: true
            },
            {
                key: 'column5',
                name: '',
                fieldName: 'authorname',
                minWidth: 100,
                maxWidth: 100,
                isRowHeader: false,
                //onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IDocument) => {
                    return <span className={styles.titleAnchorStyle}>{item.displayName} </span>;
                }
            },
            {
                key: 'column6',
                name: 'Client Reference',
                fieldName: 'clientReference',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: true,
                //isCollapsable: true,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.clientReference}</span>;

                },
                isPadded: true
            },
            {
                key: 'column7',
                name: 'Client Name',
                fieldName: 'clientName',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: true,
                //isCollapsable: true,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.clientName}</span>;

                },
                isPadded: true
            },
            {
                key: 'column8',
                name: 'Matter Reference',
                fieldName: 'matterReference',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: true,
                //isCollapsable: true,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.matterReference}</span>;

                },
                isPadded: true
            },
            {
                key: 'column9',
                name: 'Matter Key Desc',
                fieldName: 'matterKeyDesc',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: true,
                //isCollapsable: true,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.matterKeyDesc}</span>;

                },
                isPadded: true
            },
            {
                key: 'column10',
                name: 'Record Name',
                fieldName: 'recordname',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: false,
                //isCollapsable: false,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.recordName}</span>;

                },
                isPadded: true
            }
        ];
        const _checkOutColumns: IColumn[] = [
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
                minWidth: 20,
                maxWidth: 20,
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span dangerouslySetInnerHTML={{ __html: item.iconName }}></span>;
                }
            }, {
                key: 'column11',
                name: ' ',
                fieldName: 'folderpath',
                minWidth: 16,
                maxWidth: 16,
                isRowHeader: true,
                //isResizable: true,
                //isSorted: true,
                //isSortedDescending: true,
                //sortAscendingAriaLabel: 'Sorted A to Z',
                //sortDescendingAriaLabel: 'Sorted Z to A',
                //onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IDocument) => {
                    return <a href={item.folderPath} className={styles.titleAnchorStyle}><div><i title="Launch Folder" className="ms-Icon ms-Icon--OpenInNewWindow" aria-hidden="true"></i></div></a>;
                }
            },
            {
                key: 'column2',
                name: 'Name',
                fieldName: 'name',
                minWidth: 150,
                maxWidth: 150,
                isRowHeader: true,
                data: 'string',
                isPadded: true,
                onRender: (item: IDocument) => {
                    return <a href={item.href} className={styles.titleAnchorStyle}>{item.name}</a>;
                }
            },
            {
                key: 'column6',
                name: 'Client Reference',
                fieldName: 'clientReference',
                minWidth: 70,
                maxWidth: 70,
                data: 'string',
                onRender: (item: IDocument) => {
                    return <span>{item.clientReference}</span>;

                },
                isPadded: true
            },
            {
                key: 'column7',
                name: 'Client Name',
                fieldName: 'clientName',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: true,
                //isCollapsable: true,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.clientName}</span>;

                },
                isPadded: true
            },
            {
                key: 'column8',
                name: 'Matter Reference',
                fieldName: 'matterReference',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: true,
                //isCollapsable: true,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.matterReference}</span>;

                },
                isPadded: true
            },
            {
                key: 'column9',
                name: 'Matter Key Desc',
                fieldName: 'matterKeyDesc',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: true,
                //isCollapsable: true,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.matterKeyDesc}</span>;

                },
                isPadded: true
            },
            {
                key: 'column10',
                name: 'Record Name',
                fieldName: 'recordname',
                minWidth: 70,
                maxWidth: 70,
                //isResizable: false,
                //isCollapsable: false,
                data: 'string',
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.recordName}</span>;

                },
                isPadded: true
            }
        ];
        const _followColumns: IColumn[] = [
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
                minWidth: 20,
                maxWidth: 20,
                //onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span dangerouslySetInnerHTML={{ __html: item.iconName }}></span>;
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
            checkOutColumns: _checkOutColumns,
            followColumns: _followColumns,
            checkedout: [],
            followed: [],
            columns: _columns,
            recentEmail: [],
            selectionDetails: this._getSelectionDetails(),
            isModalSelection: this._selection.isModal(),
            isCompactMode: false
        };
        /*function iconfinder(iconname){
            if(iconname == 'txt'){
                return <img src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/pptx_16x1.svg" className={'DetailsListExample-documentIconImage'} />;
            }
        }*/

    }
    public render() {
        const { columns, checkOutColumns, followColumns, isCompactMode, items, selectionDetails, checkedout, followed, recentEmail } = this.state;
        thisDuplicating = this;
        var absUrl = this._context.pageContext.web.absoluteUrl;
        var absurlStartIndex = absUrl.indexOf('//');
        var absurlEndIndex = absUrl.indexOf('.');
        nameOfTenant = absUrl.substring(absurlStartIndex + 2, absurlEndIndex);
        return (
            <div className={styles.displayDocuments}>
                <div className={styles.containerMD}>
                    <div className={styles.rowMD}>
                        <div className={'ms-Grid'}>
                            <div className="search ms-Grid-row">
                                <div className={'searchBocContains ' + styles.searchBocContains}>
                                    <div className="ms-SearchBoxExample">
                                        <div className={'ms-Grid-col ms-lg12 ' + styles.searchboxtitle}>My Documents</div>
                                    </div>
                                </div>
                            </div>
                            <div className={'ms-Grid-row myRecentDocuments ' + styles.recentMD}>
                                <div className={'ms-Grid-col ms-lg12 accordion5 ' + styles.accordion5}><p className={'ms-Grid-col ms-lg11 accordiontitle5 ' + styles.accordionTitle} >My Recent Documents</p><div className={styles.accordionicon5}><i className={'ms-Grid-col ms-lg1 accordionicon5 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon5} aria-hidden="true"></i></div></div>
                                <div className={'containsMyRecentDocuments ' + styles.panel}>
                                    <MarqueeSelection selection={this._selection}>
                                        <DetailsList
                                            items={items}
                                            compact={isCompactMode}
                                            columns={columns}
                                            selectionMode={this.state.isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                                            setKey="set"
                                            layoutMode={DetailsListLayoutMode.justified}
                                            isHeaderVisible={true}
                                            selection={this._selection}
                                            selectionPreservedOnEmptyClick={true}
                                            //onItemInvoked={this._onItemInvoked}
                                            enterModalSelectionOnTouch={true}
                                        />
                                    </MarqueeSelection>
                                </div>

                            </div>
                            <div className={'ms-Grid-row myCheckedOutDocs ' + styles.recentMD}>
                                <div className={'ms-Grid-col ms-lg12 accordion6 ' + styles.accordion6}><p className={'ms-Grid-col ms-lg11 accordiontitle6 ' + styles.accordionTitle} >My CheckedOut Documents</p><div className={styles.accordionicon6}><i className={'ms-Grid-col ms-lg1 accordionicon6 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon6} aria-hidden="true"></i></div></div>
                                <div className={'containsMyCheckedOutDocuments ' + styles.panel}>
                                    <MarqueeSelection selection={this._selection}>
                                        <DetailsList
                                            items={checkedout}
                                            compact={isCompactMode}
                                            columns={checkOutColumns}
                                            selectionMode={this.state.isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                                            setKey="set"
                                            layoutMode={DetailsListLayoutMode.justified}
                                            isHeaderVisible={true}
                                            selection={this._selection}
                                            selectionPreservedOnEmptyClick={true}
                                            //onItemInvoked={this._onItemInvoked}
                                            enterModalSelectionOnTouch={true}
                                        />
                                    </MarqueeSelection>
                                </div>
                            </div>
                            <div className={'ms-Grid-row myFollowedDocs ' + styles.recentMD}>
                                <div className={'ms-Grid-col ms-lg12 accordion7 ' + styles.accordion7}><p className={'ms-Grid-col ms-lg11 accordiontitle7 ' + styles.accordionTitle} >My Following Documents<a className={styles.linkAnchorStyle} href={followedLink}><i className={"ms-Icon ms-Icon--Link " + styles.linkicon} aria-hidden="true"></i></a></p><div className={styles.accordionicon7}><i className={'ms-Grid-col ms-lg1 accordionicon7 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon7} aria-hidden="true"></i></div></div>
                                <div className={'containsMyFollowedDocuments ' + styles.panel}>
                                    <MarqueeSelection selection={this._selection}>
                                        <DetailsList
                                            items={followed}
                                            compact={isCompactMode}
                                            columns={followColumns}
                                            selectionMode={this.state.isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                                            setKey="set"
                                            layoutMode={DetailsListLayoutMode.justified}
                                            isHeaderVisible={true}
                                            selection={this._selection}
                                            selectionPreservedOnEmptyClick={true}
                                            //onItemInvoked={this._onItemInvoked}
                                            enterModalSelectionOnTouch={true}
                                        />
                                    </MarqueeSelection>
                                </div>
                            </div>
                            <div className={'ms-Grid-row myFollowedDocs ' + styles.recentMD}>
                                <div className={'ms-Grid-col ms-lg12 accordion8 ' + styles.accordion8}><p className={'ms-Grid-col ms-lg11 accordiontitle8 ' + styles.accordionTitle} >My Recent Emails</p><div className={styles.accordionicon8}><i className={'ms-Grid-col ms-lg1 accordionicon8 ms-Icon ms-Icon--ChevronRight ' + styles.accordionicon8} aria-hidden="true"></i></div></div>
                                <div className={'containsMyRecentEmails ' + styles.panel}>
                                    <MarqueeSelection selection={this._selection}>
                                        <DetailsList
                                            items={recentEmail}
                                            compact={isCompactMode}
                                            columns={checkOutColumns}
                                            selectionMode={this.state.isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                                            setKey="set"
                                            layoutMode={DetailsListLayoutMode.justified}
                                            isHeaderVisible={true}
                                            selection={this._selection}
                                            selectionPreservedOnEmptyClick={true}
                                            //onItemInvoked={this._onItemInvoked}
                                            enterModalSelectionOnTouch={true}
                                        />
                                    </MarqueeSelection></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        );
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
    public getMyRecentDocs() {
        _items = [];

        pnp.sp.web.currentUser.get().then(result => {
            var loginName = result.Title;
            $.ajax({
                //url: this._context.pageContext.web.absoluteUrl + "/_api/search/query?querytext=%27(Path:" + this._context.pageContext.web.absoluteUrl + ")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%20AND(FileExtension%3C%3Emsg)AND(FileExtension%3C%3Eeml)%27&rowlimit=50&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,Editor,SPWebUrl%27&sortlist=%27LastModifiedTime:descending%27",
                url: this._context.pageContext.web.absoluteUrl + "/_api/search/query?properties=%27SourceName:MyRecentDocs,SourceLevel:SPSiteSubscription%27&selectproperties=%27Title,Path,Filename,FileExtension,Created,Modified,Author,SPWebUrl,PTLClientReference,RefinableString02,PTLMatterReference,RefinableString03,PTLFriendlyName,LastModifiedTime%27&sortlist=%27Created:descending%27&rowlimit=" + rowLimit + "",
                method: "GET",
                headers: {
                    "accept": "application/json;odata=verbose",   //It defines the Data format
                },
                cache: false,
                success: function (myRecentDocs) {
                    console.log(myRecentDocs);
                    var htmlStrRecentDocs = '';
                    let flag: number = 0;
                    //thisDuplicating.setState({items:myRecentDocs.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results})
                    $.each(myRecentDocs.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results, function (index, value) {
                        let title: String = value.Cells.results.find(x => x.Key === 'Filename').Value;
                        let path: String = value.Cells.results.find(x => x.Key === 'Path').Value;
                        let fileExt: String = value.Cells.results.find(x => x.Key === 'FileExtension').Value;
                        var displayname = loginName;
                        var splitAuthor = displayname.split(" ");
                        var displayAuthor = splitAuthor[0].charAt(0) + splitAuthor[1].charAt(0);
                        var modified = value.Cells.results.find(x => x.Key === 'LastModifiedTime').Value;
                        var sitepagesindx = path.toLowerCase().indexOf('sitepages');
                        var Created = value.Cells.results.find(x => x.Key === 'Created').Value;
                        var clientReference = value.Cells.results.find(x => x.Key === 'PTLClientReference').Value;
                        var clientName = value.Cells.results.find(x => x.Key === 'RefinableString02').Value;
                        var matterReference = value.Cells.results.find(x => x.Key === 'PTLMatterReference').Value;
                        var matterKeyDesc = value.Cells.results.find(x => x.Key === 'RefinableString03').Value;
                        var recordName = value.Cells.results.find(x => x.Key === 'PTLFriendlyName').Value;
                        var folderPath = path.substring(0, path.lastIndexOf('/'));

                        var iconurl;
                        //if (flag < 20) {
                        //flag++;
                        var date = new Date(modified);
                        var formatCreated = formatDate(date);
                        //htmlStrRecentDocs += '<a  href="' + path + '"><div id="recentdocsbox' + index + '" class="ms-Grid-col ms-lg5 recentdocsbox ' + styles.followedbox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg7 followedAnchorlib"><div class="clientAnchorlib ' + styles.followedAnchorlib + '">' + title + '</div></div><div class="drillbox ms-Grid-col ms-lg3 drillbox ' + styles.drillbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></div></a>';
                        var fileImg = thisDuplicating._context.pageContext.web.absoluteUrl + "/SiteAssets/Logo/PTLLogo.jpg";
                        var iconElement = iconManager(fileExt, fileImg);
                        if (flag < 20) {
                            flag++;
                            //iconurl = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/" + fileExt + "_16x1.svg";
                            _items.push({
                                name: title,
                                value: title,
                                folderPath: folderPath,
                                iconName: iconElement,
                                author: displayAuthor,
                                dateModified: formatCreated,
                                //dateModifiedValue: 12,
                                displayName: loginName,
                                href: path,
                                clientReference: clientReference,
                                clientName: clientName,
                                matterReference: matterReference,
                                matterKeyDesc: matterKeyDesc,
                                recordName: recordName,
                            });
                        }

                        //}
                    });
                    function formatDate(date) {
                        var monthNames = [
                            "January", "February", "March",
                            "April", "May", "June", "July",
                            "August", "September", "October",
                            "November", "December"
                        ];

                        var day = date.getDate();
                        var monthIndex = date.getMonth();
                        var year = date.getFullYear();

                        return day + ' ' + monthNames[monthIndex] + ' ' + year;
                    }
                    thisDuplicating.setState({ items: _items });
                    //$('.containsMyRecentDocuments').append(htmlStrRecentDocs);
                    var accordion5 = document.getElementsByClassName("accordion5");
                    var i;
                    for (i = 0; i < accordion5.length; i++) {
                        accordion5[i].addEventListener("click", function () {
                            this.classList.toggle("active");
                            var panel = this.nextElementSibling;
                            if (panel.style.maxHeight) {
                                panel.style.maxHeight = null;
                                $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
                                $('.accordiontitle5').css("color", "rgba(106, 191, 52, 1)");
                                $('.accordionicon5').css("color", "gray");
                                $('.accordionicon5').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
                            } else {
                                panel.style.maxHeight = panel.scrollHeight + "px";
                                $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
                                $('.accordiontitle5, .accordionicon5').css("color", "white");
                                $('.accordionicon5').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
                            }
                        });
                    }
                    $('.accordion5').trigger('click');
                },
                error: function (data) {
                    console.log(data);
                }
            });
        });
    }
    public getFollowedDocuments() {
        _followed = [];
        var serverRelativeUrl = this._context.pageContext.web.serverRelativeUrl + "/";
        pnp.sp.web.currentUser.get().then(result => {
            console.log(result);
            var stringOfEmail = result.Email.replace(/[^a-z0-9\s]/gi, '_');
            $.ajax({
                url: this._context.pageContext.web.absoluteUrl + "/_api/social.following/my/followed(types=2)",
                method: "GET",
                headers: {
                    "accept": "application/json;odata=verbose",   //It defines the Data format
                },
                cache: false,
                success: function (IFollowed) {
                    console.log(IFollowed);

                    var htmlStrFollowedDocs = '';
                    var iconurl;
                    let flag: number = 0;
                    $.each(IFollowed.d.Followed.results, function (index, value) {
                        var indexOfChecking = value.ContentUri.indexOf(serverRelativeUrl);
                        var fileextIndx = value.ContentUri.lastIndexOf('.');
                        var fileextLgth = value.ContentUri.length;
                        var filextSubStr = value.ContentUri.substring(fileextIndx + 1, fileextLgth);
                        var fileImg = thisDuplicating._context.pageContext.web.absoluteUrl + "/SiteAssets/Logo/PTLLogo.jpg";
                        var iconElement = iconManager(filextSubStr, fileImg);
                        if (flag < 20) {
                            flag++;
                            if (indexOfChecking != -1) {
                                //htmlStrFollowedDocs += '<a  href="' + value.ContentUri + '"><div id="followedbox' + index + '" class="ms-Grid-col ms-lg5 followedbox ' + styles.followedbox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg7 followedAnchorlib"><div class="clientAnchorlib ' + styles.followedAnchorlib + '">' + value.Name + '</div></div><div class="drillbox ms-Grid-col ms-lg3 drillbox ' + styles.drillbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></div></a>';
                                //iconurl = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/" + filextSubStr + "_16x1.svg";
                                _followed.push({
                                    name: value.Name,
                                    value: value.Name,
                                    folderPath: "",
                                    iconName: iconElement,
                                    author: "",
                                    dateModified: "",
                                    displayName: "",
                                    href: value.ContentUri,
                                    clientReference: "",
                                    clientName: "",
                                    matterReference: "",
                                    matterKeyDesc: "",
                                    recordName: "",
                                });
                            }
                        }
                    });
                    followedLink = "https://" + nameOfTenant + "-my.sharepoint.com/personal/" + stringOfEmail + "/Social/FollowedContent.aspx";
                    thisDuplicating.setState({ followed: _followed });

                    eventlistener();

                    //$('.accordion3').trigger('click');
                },
                error: function (data) {
                    console.log(data);
                    followedLink = "https://" + nameOfTenant + "-my.sharepoint.com/personal/" + stringOfEmail + "/Social/FollowedContent.aspx";
                    thisDuplicating.setState({ followed: _followed });
                    eventlistener();
                }
            });
        });
        function eventlistener() {
            var accordion7 = document.getElementsByClassName("accordion7");
            var i;
            for (i = 0; i < accordion7.length; i++) {
                accordion7[i].addEventListener("click", function (e) {
                    //var targetelement = e.target || e.srcElement;
                    //console.log(targetelement);
                    if (e.srcElement.classList[1] != "ms-Icon--Link") {
                        this.classList.toggle("active");
                        var panel = this.nextElementSibling;
                        if (panel.style.maxHeight) {
                            panel.style.maxHeight = null;
                            $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
                            $('.accordiontitle7').css("color", "rgba(106, 191, 52, 1)");
                            $('.accordionicon7').css("color", "gray");
                            $('.ms-Icon--Link').css("color", "rgba(106, 191, 52, 1)");
                            $('.accordionicon7').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
                        } else {
                            panel.style.maxHeight = panel.scrollHeight + "px";
                            $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
                            $('.accordiontitle7, .accordionicon7').css("color", "white");
                            $('.ms-Icon--Link').css("color", "white");
                            $('.accordionicon7').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
                        }
                    }
                });
            }
        }
    }
    public getrecentemails() {
        _recentEmail = [];
        pnp.sp.web.currentUser.get().then(result => {
            console.log(result.LoginName);
            var absurl = this._context.pageContext.web.absoluteUrl;
            $.ajax({
                //url: this._context.pageContext.web.absoluteUrl + "/_api/search/query?querytext=%27(Path:"+this._context.pageContext.web.absoluteUrl+")%27&querytemplate=%27(AuthorOwsUser:{User.AccountName}%20OR%20EditorOwsUser:{User.AccountName})%20AND%20AND%20IsDocument:1%20AND%20-Title:OneNote_DeletedPages%20AND%20-Title:OneNote_RecycleBin%20AND(FileExtension:msg%20OR%20FileExtension:eml)%27&rowlimit=50&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,SPWebUrl%27&sortlist=%27LastModifiedTime:descending%27",
                url: absurl + "/_api/search/query?properties=%27SourceName:MyRecentEmail,SourceLevel:SPSite%27&selectproperties=%27Title,Path,Filename,FileExtension,Created,Author,SPWebUrl,PTLClientReference,RefinableString02,PTLMatterReference,RefinableString03,PTLFriendlyName%27&sortlist=%27LastModifiedTime:descending%27&rowlimit=" + rowLimit + "",
                method: "GET",
                headers: {
                    "accept": "application/json;odata=verbose",   //It defines the Data format
                },
                cache: false,
                success: function (myRecentEmails) {
                    console.log(myRecentEmails);
                    var htmlStrRecentEmails = '';
                    let flag: number = 0;
                    $.each(myRecentEmails.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results, function (index, value) {
                        let title: String = value.Cells.results.find(x => x.Key === 'Filename').Value;
                        let path: String = value.Cells.results.find(x => x.Key === 'Path').Value;
                        let fileExt: String = value.Cells.results.find(x => x.Key === 'FileExtension').Value;
                        var clientReference = value.Cells.results.find(x => x.Key === 'PTLClientReference').Value;
                        var clientName = value.Cells.results.find(x => x.Key === 'RefinableString02').Value;
                        var matterReference = value.Cells.results.find(x => x.Key === 'PTLMatterReference').Value;
                        var matterKeyDesc = value.Cells.results.find(x => x.Key === 'RefinableString03').Value;
                        var recordName = value.Cells.results.find(x => x.Key === 'PTLFriendlyName').Value;
                        var folderPath = path.substring(0, path.lastIndexOf('/'));
                        //htmlStrRecentEmails += '<a  href="' + path + '"><div id="emailbox' + index + '" class="ms-Grid-col ms-lg5 emailbox ' + styles.followedbox + '"><div class="ms-Grid-col ms-lg2 partybox ' + styles.partybox + '"><i class="ms-Icon ms-Icon--PartyLeader" aria-hidden="true"></i></div><div class="ms-Grid-col ms-lg7 followedAnchorlib"><div class="clientAnchorlib ' + styles.followedAnchorlib + '">' + title + '</div></div><div class="drillbox ms-Grid-col ms-lg3 drillbox ' + styles.drillbox + '"><i class="ms-Icon ms-Icon--DrillDown" aria-hidden="true"></i></div></div></a>';
                        var fileImg = thisDuplicating._context.pageContext.web.absoluteUrl + "/SiteAssets/Logo/PTLLogo.jpg";
                        var iconElement = iconManager(fileExt, fileImg);

                        if (flag < 20) {
                            flag++;
                            _recentEmail.push({
                                name: title,
                                value: title,
                                folderPath: folderPath,
                                iconName: iconElement,
                                author: "",
                                dateModified: "",
                                displayName: "",
                                href: path,
                                clientReference: clientReference,
                                clientName: clientName,
                                matterReference: matterReference,
                                matterKeyDesc: matterKeyDesc,
                                recordName: recordName,
                            });
                        }
                    });
                    thisDuplicating.setState({ recentEmail: _recentEmail });
                    //('.containsMyRecentEmails').append(htmlStrRecentEmails);

                    var accordion8 = document.getElementsByClassName("accordion8");
                    var i;
                    for (i = 0; i < accordion8.length; i++) {
                        accordion8[i].addEventListener("click", function () {
                            this.classList.toggle("active");
                            var panel = this.nextElementSibling;
                            if (panel.style.maxHeight) {
                                panel.style.maxHeight = null;
                                $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
                                $('.accordiontitle8').css("color", "rgba(106, 191, 52, 1)");
                                $('.accordionicon8').css("color", "gray");
                                $('.accordionicon8').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
                            } else {
                                panel.style.maxHeight = panel.scrollHeight + "px";
                                $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
                                $('.accordiontitle8, .accordionicon8').css("color", "white");
                                $('.accordionicon8').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
                            }
                        });
                    }
                },
                error: function (data) {
                    console.log(data);
                }
            });
        });
    }
    public getCheckedOutDocuments() {
        _checkedout = [];
        var serverRelativeUrl = this._context.pageContext.web.serverRelativeUrl + "/";
        $.ajax({
            url: this._context.pageContext.web.absoluteUrl + "/_api/search/query?properties='SourceName:MyCheckedOutDocs,SourceLevel:SPSite'&selectproperties='Title,Path,Filename,OriginalPath,ModifiedOWSDATE,SiteTitle,ID,SPSiteURL,FileExtension,PTLClientReference,RefinableString02,PTLMatterReference,RefinableString03,PTLFriendlyName'&rowlimit=" + rowLimit + "",
            method: "GET",
            headers: {
                "accept": "application/json;odata=verbose",   //It defines the Data format
            },
            cache: false,
            success: function (checkedoutdocs) {
                console.log(checkedoutdocs);
                let flag: number = 0;
                var iconurl;
                $.each(checkedoutdocs.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results, function (index, value) {
                    let title: String = value.Cells.results.find(x => x.Key === 'Filename').Value.toLowerCase();
                    let checkedOutDocUrl: string = value.Cells.results.find(x => x.Key === 'OriginalPath').Value;
                    let fileExt: String = value.Cells.results.find(x => x.Key === 'FileExtension').Value;
                    let path: String = value.Cells.results.find(x => x.Key === 'Path').Value;
                    var clientReference = value.Cells.results.find(x => x.Key === 'PTLClientReference').Value;
                    var clientName = value.Cells.results.find(x => x.Key === 'RefinableString02').Value;
                    var matterReference = value.Cells.results.find(x => x.Key === 'PTLMatterReference').Value;
                    var matterKeyDesc = value.Cells.results.find(x => x.Key === 'RefinableString03').Value;
                    var recordName = value.Cells.results.find(x => x.Key === 'PTLFriendlyName').Value;
                    var folderPath = path.substring(0, path.lastIndexOf('/'));
                    var fileImg = thisDuplicating._context.pageContext.web.absoluteUrl + "/SiteAssets/Logo/PTLLogo.jpg";
                    var iconElement = iconManager(fileExt, fileImg);
                    if (flag < 20) {
                        flag++;
                        _checkedout.push({
                            name: title,
                            value: title,
                            folderPath: folderPath,
                            iconName: iconElement,
                            author: "",
                            dateModified: "",
                            displayName: "",
                            href: checkedOutDocUrl,
                            clientReference: clientReference,
                            clientName: clientName,
                            matterReference: matterReference,
                            matterKeyDesc: matterKeyDesc,
                            recordName: recordName,
                        });
                        //flag++;

                    }
                });
                thisDuplicating.setState({ checkedout: _checkedout });
                var accordion6 = document.getElementsByClassName("accordion6");
                var i;
                for (i = 0; i < accordion6.length; i++) {
                    accordion6[i].addEventListener("click", function () {
                        this.classList.toggle("active");
                        var panel = this.nextElementSibling;
                        if (panel.style.maxHeight) {
                            panel.style.maxHeight = null;
                            $(this).css({ "cssText": "border: 1px solid rgba(106, 191, 52, 1) !important;", "background-color": "white" });
                            $('.accordiontitle6').css("color", "rgba(106, 191, 52, 1)");
                            $('.accordionicon6').css("color", "gray");
                            $('.accordionicon6').removeClass("ms-Icon--ChevronDown").addClass("ms-Icon--ChevronRight");
                        } else {
                            panel.style.maxHeight = panel.scrollHeight + "px";
                            $(this).css({ "cssText": "border:none !important", "background-color": "rgba(55, 55, 55, 1)" });
                            $('.accordiontitle6, .accordionicon6').css("color", "white");
                            $('.accordionicon6').removeClass("ms-Icon--ChevronRight").addClass("ms-Icon--ChevronDown");
                        }
                    });
                }
            },
            error: function (data) {
                console.log(data);
            }
        });
    }

}
function iconManager(filextSubStr, fileImg) {
    //var iconurl = "<img src="+fileImg+">";
    var iconurl;
    iconurl = "<i class='ms-Icon ms-Icon--Document' style='font-size:16px;'></i>";
    if (filextSubStr == "txt") {
        iconurl = "<i class='ms-Icon ms-Icon--TextDocument' style='font-size:16px;'></i>";
    }
    if (filextSubStr == "docx" || filextSubStr == "docm" || filextSubStr == "dotx" || filextSubStr == "dotm" || filextSubStr == "doc") {
        iconurl = "<i class='ms-Icon ms-Icon--WordDocument' style='font-size:16px;'></i>";
    }
    if (filextSubStr == "xlsx" || filextSubStr == "xlsm" || filextSubStr == "xltx" || filextSubStr == "xltm") {
        iconurl = "<i class='ms-Icon ms-Icon--ExcelDocument' style='font-size:16px;'></i>";
    }
    if (filextSubStr == "pdf") {
        iconurl = "<i class='ms-Icon ms-Icon--PDF' style='font-size:16px;'></i>";
    }
    if (filextSubStr == "pptx" || filextSubStr == "pptm" || filextSubStr == "potx" || filextSubStr == "potm" || filextSubStr == "ppam" || filextSubStr == "ppsx" || filextSubStr == "ppsm" || filextSubStr == "sldx" || filextSubStr == "sldm") {
        iconurl = "<i class='ms-Icon ms-Icon--PowerPointDocument' style='font-size:16px;'></i>";
    }
    if (filextSubStr == "eml" || filextSubStr == "msg") {
        iconurl = "<i class='ms-Icon ms-Icon--MailSolid' style='font-size:16px;'></i>";
    }
    return iconurl;
}