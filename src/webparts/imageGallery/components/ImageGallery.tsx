import * as React from 'react';
import styles from './ImageGallery.module.scss';
import { IImageGalleryProps } from './IImageGalleryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { css, classNamesFunction, IStyleFunction } from '@uifabric/utilities/lib';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IListService, IImage } from '../../../Interfaces';
import { ListService } from '../../../Services/ListService';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { objectDefinedNotNull, stringIsNullOrEmpty } from '@pnp/common';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import Paging from './Paging/Paging';

export interface IImageGalleryState {
  showPanel: boolean;
  selectedImage?: IImage;
  showLoader: boolean;
  itemsNotFoundMessage?: string,
  sQuery?: string,
  dQuery?: string
  itemsNotFound?: boolean,
  itemCount?: number;
  pageSize?: number;
  currentPage?: number;
  items?: any[];
  status?: string;
}

export default class ImageGallery extends React.Component<IImageGalleryProps, IImageGalleryState> {

  private _spService: IListService;
  private selectQuery: string[] = [];
  private expandQuery: string[] = [];
  private filterQuery: string[] = [];
  /**
   *
   */
  constructor(props: IImageGalleryProps, state: IImageGalleryState) {
    super(props);

    this.state = {
      items: [],
      showPanel: false,
      selectedImage: {} as IImage,
      showLoader: false,
      itemsNotFound: false,
      pageSize: this.props.pageSize,
      currentPage: 1,


    }



    this._onTaxPickerChange = this._onTaxPickerChange.bind(this);
    this._onPageUpdate = this._onPageUpdate.bind(this);
    this._onImageClick = this._onImageClick.bind(this);
    this._onSearchChange = this._onSearchChange.bind(this);
    this._spService = new ListService(this.props.context.spHttpClient);
  }


  public async componentDidMount() {
    //Get Images from the library 



    let value = await this._spService.getListItemsCount(`${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/ItemCount`);
    this.setState({
      itemCount: value
    });

    const queryParam = this.buildQueryParams();
    this._readItems(`${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items/${queryParam}`)

  }

  private async _readItems(url: string) {
    this.setState({
      items: [],
      status: 'Loading all items...',
      showLoader: true
    });
    let items = await this._spService.readItems(url);

    this.setState({
      showLoader: false,
      items,
      status: `Showing items ${(this.state.currentPage - 1) * this.props.pageSize + 1} - ${(this.state.currentPage - 1) * this.props.pageSize + items.length} of ${this.state.itemCount}`
    });

  }

  private _onPageUpdate(pageNumber: number) {
    //this.readItems()
    this.setState({
      currentPage: pageNumber,
    });
    const p_ID = (pageNumber - 1) * this.props.pageSize;
    const selectColumns = '&$select=' + this.selectQuery;
    const expandColumns = '&$expand=' + this.expandQuery;
    const filterColumns = '&$filter=' + this.filterQuery;
    const queryParam = `%24skiptoken=Paged%3dTRUE%26p_ID=${p_ID}&$top=${this.props.pageSize}`;
    var url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items?` + queryParam + selectColumns + filterColumns + expandColumns;
    this._readItems(url);
  }



  private buildQueryParams(taxQuery?: string, searchQuery?: string): string {
    this.selectQuery = [];
    this.expandQuery = [];
    this.filterQuery = [];

    this.selectQuery.push("ID");
    this.selectQuery.push("Title");
    this.selectQuery.push("FileRef");
    this.selectQuery.push("FileLeafRef");
    this.selectQuery.push("Department");
    this.selectQuery.push("TaxCatchAll/Term");

    this.expandQuery.push("TaxCatchAll");

    this.filterQuery.push(this.buildFilterQuery(taxQuery, searchQuery));

    const queryParam = `?%24skiptoken=Paged%3dTRUE%26p_ID=1&$top=${this.state.pageSize}`;
    const selectColumns = this.selectQuery === null || this.selectQuery === undefined || this.selectQuery.length === 0 ? "" : '&$select=' + this.selectQuery.join();
    const filterColumns = this.filterQuery === null || this.filterQuery === undefined || this.filterQuery.length === 0 ? "" : '&$filter=' + this.filterQuery.join();
    const expandColumns = this.expandQuery === null || this.expandQuery === undefined || this.expandQuery.length === 0 ? "" : '&$expand=' + this.expandQuery.join();
    return queryParam + selectColumns + filterColumns + expandColumns;
  }


  

  private buildFilterQuery(taxQuery: string, searchQuery: string) {
    let result: string = "";

    if (!stringIsNullOrEmpty(taxQuery) && stringIsNullOrEmpty(searchQuery)) {
      result = `TaxCatchAll/Term eq '${taxQuery}'`;
    }

    if (stringIsNullOrEmpty(taxQuery) && !stringIsNullOrEmpty(searchQuery)) {
      result = `startswith(Title,'${searchQuery}')`;
    }

    if (!stringIsNullOrEmpty(taxQuery) && !stringIsNullOrEmpty(searchQuery)) {
      result = `(TaxCatchAll/Term eq '${taxQuery}') and (startswith(Title,'${searchQuery}'))`;
    }
    if (stringIsNullOrEmpty(taxQuery) && stringIsNullOrEmpty(searchQuery)) {
      result = "";
    }

    return result;

  }
  private async _onTaxPickerChange(terms: IPickerTerms) {

    let query = "";

    query = terms.length && terms[0].name ? terms[0].name : "";

    this.setState({
      dQuery: query
    });
    
    let queryParam = this.buildQueryParams(query, this.state.sQuery);
    this._readItems(`${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items/${queryParam}`);
    
  }
  private async _onSearchChange(query: any) {
    this.setState({
      sQuery: query
    });
    let queryParam = this.buildQueryParams(this.state.dQuery, query);

    this._readItems(`${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items/${queryParam}`);

  }
  private _onImageClick(selectedImage: any): void {
    this.setState({
      selectedImage,
      showPanel: true
    });

  }



  public render(): React.ReactElement<IImageGalleryProps> {

    const spinnerStyles = props => ({
      circle: [
        {
          width: '60px',
          height: '60px',
          borderWidth: '4px',
          selectors: {
            ':hover': {
              background: 'f8f8ff8',
            }
          }
        }
      ]
    });


    let result = [];

    let tagList;

    if (this.state.items.length) {

      result = this.state.items.map((item, index) => {
        return (
          <div key={index} className={css(styles.column, styles.mslg3)} onClick={() => this._onImageClick(item)}>

            <div className={css(styles.thumbnail)}>
              <img src={item.FileRef} title={item.Title} id={item.Id} />
              <figcaption>{item.Title}</figcaption>
            </div>
          </div>
        )
      });
    }

    if (objectDefinedNotNull(this.state.selectedImage.Department)) {
      tagList = this.state.selectedImage.Department.map((tag: any, index) => {
        return <li className={styles.listGroupItem} key={index}> <Icon iconName="Tag" className={styles.msIconTag} /> {tag.Label}</li>;
      });
    }
    return (
      <div className={styles.imageGallery}>
        <div className={styles.container} dir="ltr">
          <div className={css(styles.row, styles.header)}>
            <div className={css(styles.column, styles.mslg12, styles.pageTitle)}>
              <h1>Image Gallery <small> Filterable</small></h1></div>

          </div>
          <div className={css(styles.row, styles.filters)}>
            <div className={css(styles.column, styles.mslg12, styles.panel)}>
              <div className={styles.panelBody}>
                <div className={css(styles.column, styles.mslg3, styles.filter)}>
                  <TaxonomyPicker
                    allowMultipleSelections={false}
                    termsetNameOrID="Departments"
                    panelTitle="Select Term"
                    label="Filter by department"
                    context={this.props.context}
                    onChange={this._onTaxPickerChange}
                    isTermSetSelectable={false}
                  />

                </div>
                <div className={css(styles.column, styles.mslg3, "ms-u-lgPush6", styles.searchBox)}>
                  <TextField label="Search" className={styles.searchBoxInputField} placeholder="Enter search term" onChanged={this._onSearchChange} />
                </div>
              </div>
            </div>
          </div>
          <div className={css(styles.row)}>
            <div className={css(styles.column, styles.mslg12, styles.panel)}>
              <div className={styles.panelBody}>


                {
                  this.state.showLoader
                    ? <Spinner size={SpinnerSize.large} label="loading..." className={css(styles.loader)} getStyles={spinnerStyles} />
                    : ""
                }

                <div className={css(styles.row, styles.mainContent)}>

                  {result.length > 0 ? result : ""}
                  {!result.length && this.state.itemsNotFound ? <MessageBar
                    messageBarType={MessageBarType.warning}
                    isMultiline={false}
                    // onDismiss={log('test')}
                    dismissButtonAriaLabel="Close"
                  >
                    Items not found. Try different search keyword
                  </MessageBar> : ""}
                  <Panel
                    isOpen={this.state.showPanel}
                    type={PanelType.medium}
                    // tslint:disable-next-line:jsx-no-lambda
                    onDismiss={() => this.setState({ showPanel: false })}
                    headerText={this.state.selectedImage.Title}
                  >
                    <div className={styles.modalContent}>
                      <div className={styles.modalBody}>
                        <div className={styles.thumbnail}>
                          <img src={this.state.selectedImage.FileRef} title={this.state.selectedImage.Title} id={this.state.selectedImage.Id} />
                        </div>
                        <h3>Tags</h3>
                        {this.state.selectedImage.Department ?
                          <ul className={styles.listGroup}>
                            {
                              tagList
                            }
                          </ul> : ""}
                      </div>
                    </div>
                  </Panel>

                </div>
              </div>
            </div>
          </div>
          <div className={css(styles.row, styles.pagination)}>
            <div className={css(styles.column, styles.mslg12, styles.panel)}>

              <div className={styles.panelBody}>
                <ul className={styles.pager}>

                  <div className={styles.status}>{this.state.status}</div>
                  <Paging
                    totalItems={this.state.itemCount}
                    itemsCountPerPage={this.state.pageSize}
                    onPageUpdate={this._onPageUpdate}
                    currentPage={this.state.currentPage} />
                </ul>
              </div></div>
          </div>
        </div>
      </div>
    );
  }

}
