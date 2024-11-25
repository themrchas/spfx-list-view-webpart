import '@pnp/polyfill-ie11'; // for Promises
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneTextField }
  from '@microsoft/sp-webpart-base';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
// import { sp } from '@pnp/sp';
// note we are getting the sp variable from this library, it extends
// the sp export from @pnp/sp to add the required helper methods
import { sp, SPRequestExecutorClient } from '@pnp/sp-addinhelpers';
import { Web } from '@pnp/sp';
import { Logger, ConsoleListener, LogLevel } from '@pnp/logging';
import styles from './BrandingItemViewWebPart.module.scss';
import * as strings from 'BrandingItemViewWebPartStrings';
import MockSPClient from './MockSPClient';
import LiveSPClient from './LiveSPClient';

export interface IBrandingItemViewWebPartProps {
  title: string;
  titlecolor: string;
  aligntitle: string;
  imageurl: string;
  headerheight: string;
  headerbackgroundcolor: string;
  borderthickness: string;
  bordercolor: string;
  headerborderthickness: string;
  headerbordercolor: string;
  partheight: string;
  weburl: string;
  listname: string;
  viewname: string;
  fieldname: string;
  aligncontent: string;
  separator: string;
  separatortext: string;
  showmoretext: string;
  showmoreurl: string;
}
export interface ISPList {
  Title: string;
  Id: string;
}
export interface ISPView {
  Title: string;
  Id: string;
}
export interface ISPField {
  DisplayName: string;
  InternalName: string;
}
export interface IDataSource {
  setWeb(web: Web): void;
  setList(listName: string): void;
  setView(viewName: string): void;
  setField(internalName: string): void;
  getLists(): Promise<ISPList[]>;
  getViews(): Promise<ISPView[]>;
  getFields(): Promise<ISPField[]>;
  getContents(): Promise<string[]>;
}

export default class BrandingItemViewWebPart extends BaseClientSideWebPart<IBrandingItemViewWebPartProps> {
  private dataSource: IDataSource;
  private propertyPaneDataLoaded: boolean;
  private listDropDownOptions: IPropertyPaneDropdownOption[];
  private viewDropDownOptions: IPropertyPaneDropdownOption[];
  private fieldDropDownOptions: IPropertyPaneDropdownOption[];
  private alignDropDownOptions: IPropertyPaneDropdownOption[];

  constructor() {
    super();
    this.propertyPaneDataLoaded = false;
    this.listDropDownOptions = [];
    this.viewDropDownOptions = [];
    this.fieldDropDownOptions = [];
    this.alignDropDownOptions = [];
  }

  protected onInit(): Promise<void> {
    // subscribe a listener
    Logger.subscribe(new ConsoleListener());
    // set the active log level
    Logger.activeLogLevel = LogLevel.Info;

    // Logger.writeJSON(this.properties);

    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      // Logger.write('running in workbench');
      this.dataSource = new MockSPClient();
    } else if (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    ) {
      // Logger.write('running in SharePoint');
      this.dataSource = new LiveSPClient();
    }
    this.pushWeb();
    this.dataSource.setList(this.properties.listname);
    this.dataSource.setView(this.properties.viewname);
    this.dataSource.setField(this.properties.fieldname);
    return super.onInit()
      .then(_ => {
        sp.setup({
          spfxContext: this.context,
          sp: {
            // for cross-domain queries
            fetchClientFactory: () => {
              return new SPRequestExecutorClient();
            }
          }
        });
      });
  }

  protected pushWeb(): void {
    if (this.properties.weburl) {
      // do  we need to use the crossDomainWeb method to make our requests to the host web?
      const addInWebUrl: string = this.properties.weburl;
      const hostWebUrl: string = this.context.pageContext.web.absoluteUrl;
      if ((new URL(addInWebUrl).hostname.toLowerCase()) !== (new URL(hostWebUrl).hostname.toLowerCase())) {
        // cross domain
        // make requests into the host web via the SP.RequestExecutor
        Logger.write(`Connecting to configured web using sp.crossDomainWeb(${addInWebUrl}, ${hostWebUrl})`);
        this.dataSource.setWeb(sp.crossDomainWeb(addInWebUrl, hostWebUrl));
      } else {
        // same domain
        Logger.write(`Connecting to configured web using new Web(${this.properties.weburl})`);
        this.dataSource.setWeb(new Web(this.properties.weburl));
      }
    } else {
      // no weburl specified, use current site
      Logger.write(`Connecting to current web using new Web(${this.context.pageContext.web.absoluteUrl})`);
      this.dataSource.setWeb(new Web(this.context.pageContext.web.absoluteUrl));
    }
  }

  protected _renderContent(): void {
    // Logger.write('_renderContent() entry');
    let html: string = '';
    const listContainer: Element = this.domElement.querySelector('#contents');
    this.dataSource.getContents()
      .then((contents: string[]): void => {
        let startTag: string = '<' + this.properties.separator + '>';
        let endTag: string = '</' + this.properties.separator + '>';
        let skipLast: boolean = false;
        if (this.properties.separator === 'br') {
          // breaks have only a single closed tag
          startTag = '';
          endTag = '<br/>';
          skipLast = true;
        }
        const separator: string = this.properties.separator === 'span' ? this.properties.separatortext : '';
        for (let i: number = 0; i < contents.length; i++) {
          const content: string = contents[i];
          if (content) {
            html = html.concat(
              i === 0 ? '' : separator,
              startTag,
              content,
              skipLast && i === contents.length - 1 ? '' : endTag);
          }
        }
        if (html !== '') {
          listContainer.innerHTML = html;
        }
      })
      .catch((err) => {
        html = 'Error: '.concat(escape(err.message)); // something bad happened
        // tslint:disable-next-line:no-any
      }).then((): any => {
        listContainer.innerHTML = html;
        // Logger.write('_renderContent() html pushed');
      });
  }
  protected ensureSizeStr(s: string): string {
    const parsed: number = parseInt(s, 10);
    if (isNaN(parsed)) {
      // apparently the user put in a unit
      return s;
    }
    return s + 'px'; // s was a number, add px as default units
  }
  public render(): void {
    let partStyle: string = '';
    let headerStyle: string = '';
    let titleStyle: string = '';
    let image: string = '';
    let contentStyle: string = '';
    let showMore: string = '';

    if (this.properties.bordercolor !== '' || this.properties.borderthickness !== '') {
      partStyle += `border-style: solid; `;
      if (this.properties.bordercolor !== '') {
        partStyle += `border-color: ${this.properties.bordercolor}; `;
      }
      if (this.properties.borderthickness !== '') {
        partStyle += `border-width: ${this.ensureSizeStr(this.properties.borderthickness)}; `;
      }
    }
    if (this.properties.partheight) {
      partStyle += `min-height: ${this.ensureSizeStr(this.properties.partheight)}; `;
    }
    // make full attribute string if needed
    if (partStyle !== '') { partStyle = 'style=\'' + partStyle + '\''; }

    if (this.properties.titlecolor !== '') {
      titleStyle += `color: ${this.properties.titlecolor}; `;
    }
    if (this.properties.aligntitle !== '') {
      titleStyle += `text-align: ${this.properties.aligntitle}; `;
    }
    // make full attribute string if needed
    if (titleStyle !== '') { titleStyle = 'style=\'' + titleStyle + '\''; }

    if (this.properties.headerheight !== '') {
      headerStyle += `min-height: ${this.ensureSizeStr(this.properties.headerheight)}; `;
    }
    if (this.properties.headerbackgroundcolor !== '') {
      headerStyle += `background-color: ${this.properties.headerbackgroundcolor}; `;
    }
    if (this.properties.headerbordercolor !== '' || this.properties.headerborderthickness !== '') {
      headerStyle += `border-bottom-style: solid; `;
      if (this.properties.headerbordercolor !== '') {
        headerStyle += `border-bottom-color: ${this.properties.headerbordercolor}; `;
      }
      if (this.properties.headerborderthickness !== '') {
        headerStyle += `border-bottom-width: ${this.ensureSizeStr(this.properties.headerborderthickness)}; `;
      }
    }
    // make full attribute string if needed
    if (headerStyle !== '') { headerStyle = 'style=\'' + headerStyle + '\''; }

    if (this.properties.imageurl) {
      image = `<img class='${styles.headerimg}' src='${this.properties.imageurl}' alt='[]'/>`;
    }

    if (this.properties.aligncontent !== '') {
      contentStyle += `text-align: ${this.properties.aligncontent}; `;
    }
    // make full attribute string if needed
    if (contentStyle !== '') { contentStyle = 'style=\'' + contentStyle + '\''; }

    if (this.properties.showmoretext !== '' && this.properties.showmoreurl !== '') {
      // tslint:disable-next-line:max-line-length
      showMore = `<div class='${styles.footer}'><a class='${styles.showmore}' href='${this.properties.showmoreurl}'>${this.properties.showmoretext}</a></div>`;
    }
    this.domElement.innerHTML = `
      <div class='${styles.brandingItemView}'>
        <div ${partStyle}>
          <div class='${styles.header}' ${headerStyle}>
            <p class='${styles.title}' ${titleStyle}>${escape(this.properties.title)}</p>
            ${image}
          </div>
          <div class='${styles.bottom}'>
            <div class='${styles.contentContainer}'>
              <div class='${styles.content}' ${contentStyle} id='contents'></div>
            </div>
            ${showMore}
          </div>
        </div>
      </div>`;
    this._renderContent();
  }

  // @ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('1.0.0');
  }

  protected LoadPropertyPaneData(): void {
    // Logger.write('LoadPropertyPaneData() - entry');
    if (this.alignDropDownOptions !== []) {
      this.alignDropDownOptions.push({ key: 'left', text: 'Left' });
      this.alignDropDownOptions.push({ key: 'center', text: 'Center' });
      this.alignDropDownOptions.push({ key: 'right', text: 'Right' });
      this.alignDropDownOptions.push({ key: 'justify', text: 'Justify' });
    }
    // Loads list names, views, and fields as required.
    // This is important for the first load to have the
    // currently configured data selected in the controls.
    // Combine promise waits to have a single refresh of the properties pane
    const p0: Promise<boolean> = this.loadLists();
    const p1: Promise<boolean> = this.loadViews();
    const p2: Promise<boolean> = this.loadFields();
    Promise.all([p0, p1, p2])
      .then(() => {
        // Logger.write('LoadPropertyPaneData() - lists, views and fields loaded');
        this.context.propertyPane.refresh();
      });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Logger.write('getPropertyPaneConfiguration()');
    if (!this.propertyPaneDataLoaded) {
      this.propertyPaneDataLoaded = true;
      this.LoadPropertyPaneData();
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleLabel
                }),
                PropertyPaneTextField('titlecolor', {
                  label: strings.TitleColor
                }),
                PropertyPaneDropdown('aligntitle', {
                  label: strings.AlignTitle,
                  selectedKey: this.properties.aligncontent,
                  options: this.alignDropDownOptions
                }),
                PropertyPaneTextField('imageurl', {
                  label: strings.ImageUrlLabel
                }),
                PropertyPaneTextField('headerbackgroundcolor', {
                  label: strings.HeaderBackgroundColorLabel
                }),
                PropertyPaneTextField('headerheight', {
                  label: strings.HeaderHeight
                }),
                PropertyPaneTextField('borderthickness', {
                  label: strings.BorderThicknessLabel
                }),
                PropertyPaneTextField('bordercolor', {
                  label: strings.BorderColorLabel
                }),
                PropertyPaneTextField('headerborderthickness', {
                  label: strings.HeaderBorderThicknessLabel
                }),
                PropertyPaneTextField('headerbordercolor', {
                  label: strings.HeaderBorderColorLabel
                }),
                PropertyPaneTextField('partheight', {
                  label: strings.PartHeight
                }),
                PropertyPaneDropdown('aligncontent', {
                  label: strings.AlignContent,
                  selectedKey: this.properties.aligncontent,
                  options: this.alignDropDownOptions
                }),
                PropertyPaneDropdown('separator', {
                  label: strings.SeparatorLabel,
                  selectedKey: this.properties.separator,
                  options: [
                    { key: 'p', text: 'Paragraphs' },
                    { key: 'br', text: 'Line Breaks' },
                    { key: 'span', text: 'Custom text' }
                  ]
                }),
                PropertyPaneTextField('separatortext', {
                  label: strings.SeparatorTextLabel
                })
              ]
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('weburl', {
                  label: strings.WebUrlLabel
                }),
                PropertyPaneDropdown('listname', {
                  label: strings.ListNameLabel,
                  selectedKey: this.properties.listname,
                  options: this.listDropDownOptions,
                  disabled: this.listDropDownOptions.length === 0
                }),
                PropertyPaneDropdown('viewname', {
                  label: strings.ViewNameLabel,
                  selectedKey: this.properties.viewname,
                  options: this.viewDropDownOptions,
                  disabled: this.viewDropDownOptions.length === 0
                }),
                PropertyPaneDropdown('fieldname', {
                  label: strings.FieldNameLabel,
                  selectedKey: this.properties.fieldname,
                  options: this.fieldDropDownOptions,
                  disabled: this.fieldDropDownOptions.length === 0
                }),
                PropertyPaneTextField('showmoretext', {
                  label: strings.ShowMoreTextLabel
                }),
                PropertyPaneTextField('showmoreurl', {
                  label: strings.ShowMoreUrlLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected loadLists(): Promise<boolean> {
    // Logger.write('loadLists() - entry');
    if (this.listDropDownOptions.length > 0) {
      // Logger.write('loadLists() - already loaded');
      return new Promise((resolve) => { resolve(true); });
    } else {
      // add the current value as minimum option
      if (this.properties.listname + '' !== '') {
        // Logger.write('loadLists() - add current value');
        this.listDropDownOptions.push({
          key: this.properties.listname,
          text: this.properties.listname
        });
      }
      // Logger.write('loadLists() - reading actual data');
      return new Promise<boolean>((resolve) => {
        this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Reading lists');
        this.dataSource.getLists()
          .then((data: ISPList[]) => {
            // Logger.write('loadLists() - received actual data');
            this.listDropDownOptions = [];
            data.forEach((l: ISPList) => {
              const option: IPropertyPaneDropdownOption = {
                key: l.Title,
                text: l.Title
              };
              this.listDropDownOptions.push(option);
            });
            this.context.statusRenderer.clearLoadingIndicator(this.domElement);
            // Logger.write('loadLists() - exit');
            return resolve(true);
          })
          .catch(() => {
            return resolve(false);
          });
      });
    }
  }

  protected loadViews(): Promise<boolean> {
    // Logger.write('loadViews() - entry');
    if (this.viewDropDownOptions.length > 0) {
      return new Promise((resolve) => { resolve(true); });
    } else {
      // clear list
      this.viewDropDownOptions = [];
      // add the current value as minimum option
      if (this.properties.viewname + '' !== '') {
        this.viewDropDownOptions.push({
          key: this.properties.viewname,
          text: this.properties.viewname
        });
      }
      return new Promise<boolean>((resolve) => {
        if (this.properties.listname) {
          this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Reading views from list');
          this.dataSource.getViews()
            .then((data: ISPView[]) => {
              this.viewDropDownOptions = [];
              data.forEach((v: ISPView) => {
                this.viewDropDownOptions.push({ key: v.Title, text: v.Title });
              });
              this.context.statusRenderer.clearLoadingIndicator(this.domElement);
              // Logger.write('loadViews() - exit');
              return resolve(true);
            })
            .catch(() => {
              return resolve(false);
            });
        } else {
          return resolve(false);
        }
      });
    }
  }

  protected loadFields(): Promise<boolean> {
    // Logger.write('loadFields() - entry');
    if (this.fieldDropDownOptions.length > 0) {
      return new Promise((resolve) => { resolve(true); });
    } else {
      // clear list
      this.fieldDropDownOptions = [];
      // add the current value as minimum option
      if (this.properties.viewname + '' !== '') {
        this.fieldDropDownOptions.push({
          key: this.properties.fieldname,
          text: this.properties.fieldname
        });
      }
      return new Promise<boolean>((resolve) => {
        if (this.properties.listname) {
          this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Reading fields from list');
          this.dataSource.getFields()
            .then((data: ISPField[]) => {
              this.fieldDropDownOptions = [];
              data.forEach((f: ISPField) => {
                this.fieldDropDownOptions.push({ key: f.InternalName, text: f.DisplayName });
              });
              this.context.statusRenderer.clearLoadingIndicator(this.domElement);
              // Logger.write('loadFields() - exit');
              return resolve(true);
            })
            .catch(() => {
              return resolve(false);
            });
        } else {
          return resolve(false);
        }
      });
    }
  }

  // tslint:disable-next-line:no-any
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    Logger.write('onPropertyPaneFieldChanged(\'' + propertyPath + '\', \'' + oldValue + '\', \'' + newValue + '\')');
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'weburl') {
      if (oldValue !== newValue) {
        // Logger.write('Switching to web: ' + this.properties.weburl);
        this.pushWeb();
        this.listDropDownOptions = [];
        // read lists from this web.
        this.loadLists()
          .then((): void => {
            // Logger.write('calling this.context.propertyPane.refresh()');
            this.context.propertyPane.refresh();
          });
      }
    } else if (propertyPath === 'listname') {
      if (oldValue !== newValue) {
        this.dataSource.setList(newValue);
        this.viewDropDownOptions = [];
        this.fieldDropDownOptions = [];
        // read views and fields from this list
        Promise.all<boolean, boolean>([this.loadViews(), this.loadFields()])
          .then((): void => {
            // Logger.write('calling this.context.propertyPane.refresh()');
            this.context.propertyPane.refresh();
          });
      }
    } else if (propertyPath === 'viewname') {
      if (oldValue !== newValue) {
        this.dataSource.setView(newValue);
      }
    } else if (propertyPath === 'fieldname') {
      if (oldValue !== newValue) {
        this.dataSource.setField(newValue);
      }
    }
  }
}