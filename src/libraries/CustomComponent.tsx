import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PageContext } from '@microsoft/sp-page-context';
import {MSGraphClientFactory, SPHttpClient} from "@microsoft/sp-http";
import { initializeFileTypeIcons, getFileTypeIconProps } from '@uifabric/file-type-icons';
import { Icon, CommandBarButton, IconButton, TooltipHost } from 'office-ui-fabric-react';
import { followDocument, unFollowDocument, getFollowed, isUserManage, deleteForm } from './Services/Requests';
import ListControls from './components/ListControls/ListControls';
import styles from './customComponent.module.scss';
import toast, { Toaster } from 'react-hot-toast';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {
    pageUrlParam? : string;
    pageTitleParam? : string;
    pageFileTypeParam? : string;   
    pageId?: string;
    pageContext?: PageContext; 
    sphttpClient?: SPHttpClient;
    msGraphClientFactory?: MSGraphClientFactory;
    pages?: any;
}

export function CustomComponent (props: ICustomComponentProps){

    console.log("props.pages", props.pages);

    initializeFileTypeIcons();
    const dateOptions: any = { year: 'numeric', month: 'long', day: 'numeric' };
    
    const [myFollowedItems, setMyfollowedItems] = React.useState([]);
    const [iFrameVisible, setIFrameVisible] = React.useState(false);
    const [iFrameUrl, setIFrameUrl] = React.useState(null);
    const [editControlsVisible, setEditControlsVisible] = React.useState(false);

    React.useEffect(()=>{
        getFollowed(props.msGraphClientFactory).then(res => {
            console.log("setMyfollowedItems(res)", res);
            setMyfollowedItems(res);
        });
    }, []);
    React.useEffect(()=>{
    }, [myFollowedItems.toString()]);


    // Follow & Unfollow
    const followDocHandler = (page: any) => {
        console.log("followDocHandler", page);
        followDocument(props.msGraphClientFactory, page.SiteId, page.WebId, page.ListId, page.ListItemID).then(() => {
            setMyfollowedItems(prev => {
                const currentFollowedItems = [...prev];
                currentFollowedItems.push({name: decodeURI(page.Filename), driveId: page.DriveId});
                console.log("currentFollowedItems", currentFollowedItems);
                return currentFollowedItems;
            });
            toast.custom((t) => (
                <div className={styles.toastMsg}>
                  <Icon iconName='Accept' /> Added to <a target='_blank' href="https://www.office.com/mycontent">Favorites!</a>
                </div>
            ));
        });
    };
    const unFollowDocHandler = (page: any) => {
        console.log("followDocHandler", page);
        unFollowDocument(props.msGraphClientFactory, page.SiteId, page.WebId, page.ListId, page.ListItemID).then(()=>{
            setMyfollowedItems(prev => {
                const currentFollowedItems = prev.filter(item => !(item.name === decodeURI(page.Filename) && item.driveId === page.DriveId));
                console.log("currentFollowedItems", currentFollowedItems);
                return currentFollowedItems;
            });
            toast.custom((t) => (
                <div className={styles.toastMsg}>
                  <Icon iconName='Accept' /> Removed from <a target='_blank' href="https://www.office.com/mycontent">Favorites!</a>
                </div>
            ));
        });
    };

    // Upload, Add & View Controls
    const uploadDocumentHandler = () => {
        const docUrl = props.pages.items[0].Path;
        setIFrameUrl(`${docUrl.substring(0, docUrl.lastIndexOf('/'))}/Forms/Upload.aspx`);
        setIFrameVisible(true);
        toast.custom((t) => (
            <div className={styles.toastMsg}>
              <Icon iconName='Accept' /> Item has been added to the library. Please allow few minutes for it to update.
            </div>
        ));
    };
    const addLinkHandler = () => {
        const pageItem = props.pages.items[0];
        const listUrl = pageItem.Path.substring(0, pageItem.Path.lastIndexOf('/'));
        let encodedUrl = `List=%7B${pageItem.ListId}%7D&RootFolder=${listUrl}&ContentTypeId=${pageItem.ContentTypeId}&Source=${listUrl}/Forms/allItems.aspx`;
        encodedUrl.replace(/\//g, "%2F").replace(/:/g,"%3A");
        setIFrameUrl(`${pageItem.SPSiteUrl}/_layouts/15/NewLink.aspx?${encodedUrl}&isDlg=1`);
        setIFrameVisible(true);
        // toast.custom((t) => (
        //     <div className={styles.toastMsg}>
        //       <Icon iconName='Accept' /> Item has been added to the library. Please allow few minutes for it to update.
        //     </div>
        // ));
    };
    const viewAllHandler = () => {
        const docUrl = props.pages.items[0].Path;
        window.open(`${docUrl.substring(0, docUrl.lastIndexOf('/'))}/Forms/Allitems.aspx`, '_blank');
    };

    // Edit & Delete Controls 
    const toggleEditControls = () => {
        setEditControlsVisible(prev => !prev);
    };
    const onEditFormClickHandler = (page: any) => {        
        setIFrameUrl(`${page.Path.substring(0, page.Path.lastIndexOf('/'))}/Forms/EditForm.aspx?ID=${page.ListItemID}&isDlg=1`);
        setIFrameVisible(true);
        // toast.custom((t) => (
        //     <div className={styles.toastMsg}>
        //       <Icon iconName='Accept' /> Item has been edited. Please allow few minutes for it to update.
        //     </div>
        // ));
    };
    const onDeleteIconClick = (page: any) => {
        const urlArr = page.Path.split('/');
        const listName = urlArr[urlArr.length -2];
        deleteForm(props.sphttpClient, page.SPSiteUrl, listName, page.ListItemID);
        toast.custom((t) => (
            <div className={styles.toastMsg}>
              <Icon iconName='Accept' /> Item has been deleted from the library. Please allow few minutes for it to update.
            </div>
        ));
    };

    return(
        <>
            <Toaster position='bottom-center' toastOptions={{custom:{duration: 4000}}}/>
            <div className={styles.listViewNoWrap}>
				<table className={styles.customTable} cellPadding='0' cellSpacing='0'>
                    <colgroup>
                        <col width={'10%'} />
                        {editControlsVisible && <col width={'10%'} />}
                        <col width={'60%'} />
                        <col width={'20%'} />
                    </colgroup>
					<thead>
						<tr>
							<th></th>
							{editControlsVisible && <th></th>}
							<th>Form</th>							
							<th>Team</th>							
						</tr>
					</thead>
					<tbody>
                        {props.pages.items.map(page => {
                            return (
                                <tr key={page.ListItemID}>
                                    <td>
                                        <div className={styles.formItem}>
                                            <div className={styles.favIconBtns}>
                                                {myFollowedItems.find(item => item.name === decodeURI(page.Filename) && item.driveId === page.DriveId ) ? 
                                                    <IconButton title='Unfavorite' onClick={() => unFollowDocHandler(page)} iconProps={{iconName : 'FavoriteStarFill'}} />
                                                : 
                                                    <IconButton title='Favorite' onClick={() => followDocHandler(page)} iconProps={{iconName : 'FavoriteStar'}} />
                                                }
                                            </div>
                                            <div className={styles.cellDiv}> 
                                                {page.FileType !== 'SharePoint.Link' &&
                                                    <a className={styles.attachmentLinkDownload} href={`${page.Path}`} title='Download' download>
                                                        <Icon iconName='Download' />
                                                    </a>
                                                }                                             
                                            </div>
                                        </div>
                                    </td>
                                    {editControlsVisible &&
                                        <td>
                                            <div className={styles.editControls}>
                                                <CommandBarButton alt='Edit' onClick={()=> onEditFormClickHandler(page)} iconProps={{ iconName: 'Edit' }} />
                                                <CommandBarButton alt='Delete' onClick={() => onDeleteIconClick(page)} iconProps={{ iconName: 'Delete' }} />
                                            </div>
                                        </td>
                                    }
                                    <td>
                                        <div className={styles.formItem}>
                                            <div className={styles.cellDiv}>
                                                <TooltipHost content={`${page.FileType} file`}>
                                                    <Icon {...getFileTypeIconProps({extension: page.FileType, size: 16}) }/>
                                                </TooltipHost> 
                                                <a className={styles.defautlLink + ' ' + styles.docLink} target="_blank" data-interception="off" href={page.Path}>{page.Title}</a>
                                            </div>
                                        </div>
                                    </td>
                                    <td>
                                        {page.MMIntranetDeptSubDeptGrouping && page.MMIntranetDeptSubDeptGrouping.split('|')[1]}
                                    </td>
                                </tr>
                            );
                        })}
                    </tbody>
				</table>
			</div>


            {/* <ul className='template--defaultList'>
                {props.pages.items.map(page => {
                    return (
                        <li className='template--listItem'>
                            <div className='template--listItem--result'>
                                <div className='template--listItem--icon'>
                                    <Icon {...getFileTypeIconProps({extension: page.FileType, size: 16}) }/>
                                </div>
                                <div className='template--listItem--contentContainer'>
                                    <span className='template--listItem--title example-themePrimary'>
                                        {myFollowedItems.find(item => item.name === decodeURI(page.Filename) && item.driveId === page.DriveId ) ? 
                                            <IconButton title='Unfavorite' onClick={() => unFollowDocHandler(page)} iconProps={{iconName : 'FavoriteStarFill'}} />
                                        : 
                                            <IconButton title='Favorite' onClick={() => followDocHandler(page)} iconProps={{iconName : 'FavoriteStar'}} />
                                        }
                                        <a data-interception="off" target='_blank' href={page.Path} className='page-link'>
                                            <span>{page.Title}</span>
                                        </a>
                                    </span>
                                    <span>
                                        <p>{page.HitHighlightedSummary}</p>
                                        {page.AuthorOWSUSER &&<span className='template--listItem--author'>{page.AuthorOWSUSER.split('|')[1]}</span>}
                                        <span className='template--listItem--date'>{new Date(page.Created).toLocaleDateString('en-us', dateOptions)}</span>
                                    </span>
                                </div> 
                                <div>
                                    {page.FileType !== 'SharePoint.Link' &&
                                        <a className={styles.attachmentLinkDownload} href={`${page.Path}`} title='Download' download>
                                            <Icon iconName='Download' />
                                        </a>
                                    }
                                    {editControlsVisible && 
                                        <div className={styles.editControls}>
                                            <CommandBarButton alt='Edit' onClick={()=> onEditFormClickHandler(page)} iconProps={{ iconName: 'Edit' }} />
                                            <CommandBarButton alt='Delete' onClick={() => onDeleteIconClick(page)} iconProps={{ iconName: 'Delete' }} />
                                        </div>
                                    }
                                </div>
                            </div>
                            <div className='template--listItem--thumbnailContainer'>
                                <div className='thumbnail--image'>
                                    <img width="120" src={page.AutoPreviewImageUrl} />
                                </div>
                            </div>
                        </li>
                    );
                })}
            </ul> */}

            {isUserManage(props.pageContext) &&
                <ListControls 
                    iFrameVisible = {iFrameVisible}
                    setIFrameVisible = {setIFrameVisible}
                    iFrameUrl = {iFrameUrl}
                    addLinkHandler={addLinkHandler}
                    uploadDocumentHandler={uploadDocumentHandler}
                    toggleEditControls={toggleEditControls}
                    viewAllHandler={viewAllHandler}    
                />
            }
        </>
    );

}

export class MyCustomComponentWebComponent extends BaseWebComponent {
    
    private sphttpClient: SPHttpClient;
    private pageContext: PageContext;
    private msGraphClientFactory: MSGraphClientFactory;

    public constructor() {
        super(); 
        this._serviceScope.whenFinished(()=>{
            this.pageContext = this._serviceScope.consume(PageContext.serviceKey);
            this.sphttpClient = this._serviceScope.consume(SPHttpClient.serviceKey);
            this.msGraphClientFactory = this._serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
    }
 
    public async connectedCallback() {
        let props = this.resolveAttributes();
        const customComponent = <CustomComponent pageContext={this.pageContext} sphttpClient={this.sphttpClient} msGraphClientFactory={this.msGraphClientFactory} {...props}/>;
        ReactDOM.render(customComponent, this);
    }    
}