import { Web, List, FieldTypes } from '@pnp/sp';
import { Promise } from 'es6-promise';
import { Logger, LogLevel } from '@pnp/logging';
import { ISPList, ISPView, ISPField, IDataSource } from './BrandingItemViewWebPart';

// duplicate FieldInfo from @pnp/sp v2
interface IFieldInfo {
    Description: string;
    Hidden: boolean;
    Id: string;
    InternalName: string;
    Title: string;
    FieldTypeKind: FieldTypes;
    OutputType: FieldTypes;
    // and a lot more
}

export default class LiveSPClient implements IDataSource {
    private web: Web;
    private lists: ISPList[];
    private list: List;
    private views: ISPView[];
    private viewName: string;
    private viewQuery: string;
    private fields: ISPField[];
    private fieldInternalName: string;
    private contents: string[];

    public setWeb(web: Web): void {
        Logger.writeJSON(web);
        this.web = web;
        this.lists = undefined;
        this.list = undefined;
        this.views = undefined;
        // do not clear view name, it may still be valid
        this.viewQuery = undefined;
        this.fields = undefined;
        // do not clear field name, it may still be valid
        this.contents = undefined;
    }
    public setList(listName: string): void {
        Logger.write('LiveSPClient.setList(' + listName + ')');
        if (listName) {
            this.list = this.web.lists.getByTitle(listName);
        } else {
            this.list = undefined;
        }
        this.views = undefined;
        // do not clear view name, it may still be valid
        this.viewQuery = undefined;
        this.fields = undefined;
        // do not clear field name, it may still be valid
        this.contents = undefined;
    }
    public setView(viewName: string): void {
        Logger.write('LiveSPClient.setView(' + viewName + ')');
        this.viewName = viewName;
        this.viewQuery = undefined;
        this.contents = undefined;
    }
    public setField(internalName: string): void {
        Logger.write('LiveSPClient.setField(' + internalName + ')');
        this.fieldInternalName = internalName;
        this.contents = undefined;
    }
    public getLists(): Promise<ISPList[]> {
        Logger.write('LiveSPClient.getLists()');
        return new Promise<ISPList[]>((resolve) => {
            if (this.lists) {
                resolve(this.lists);
            } else {
                Logger.write('LiveSPClient.getLists() - reading lists');
                this.lists = [];
                this.web.lists
                    .filter('Hidden eq false')
                    .select('Title', 'Id')
                    .orderBy('Title')
                    .get()
                    .then((spqueryresults) => {
                        Logger.writeJSON(spqueryresults);
                        this.lists = spqueryresults; // these two match on Title and Id
                        resolve(this.lists);
                    })
                    .catch((err) => {
                        Logger.write('Error reading lists:', LogLevel.Error);
                        Logger.writeJSON(err, LogLevel.Error);
                        resolve(this.lists);
                    });
            }
        });
    }
    public getViews(): Promise<ISPView[]> {
        Logger.write('LiveSPClient.getViews()');
        return new Promise<ISPView[]>((resolve) => {
            if (this.views) {
                resolve(this.views);
            } else {
                // Logger.write('LiveSPClient.getLists() - reading views');
                this.views = [];
                if (this.list) {
                    this.list.views
                        .select('Id', 'Title', 'Hidden')
                        .orderBy('Title')
                        .get()
                        .then((spqueryresults) => {
                            Logger.writeJSON(spqueryresults);
                            // this.views = spqueryresults;
                            // these two match on Title and Id, but we need to filter out hidden views
                            spqueryresults.forEach((viewinfo) => {
                                // If there is no title, it is probably hidden
                                if (viewinfo.Title + '' !== '') {
                                    this.views.push(viewinfo);
                                }
                            });
                            resolve(this.views);
                        })
                        .catch((err) => {
                            Logger.write('Error reading views:', LogLevel.Error);
                            Logger.writeJSON(err, LogLevel.Error);
                            resolve(this.lists);
                        });
                } else {
                    resolve(this.views);
                }
            }
        });
    }
    protected getViewQuery(): Promise<string> {
        Logger.write('LiveSPClient.getViewQuery()');
        return new Promise<string>((resolve) => {
            if (this.viewQuery) {
                resolve(this.viewQuery);
            } else if (this.list && this.viewName) {
                // escape single quote
                this.list.views.getByTitle(this.viewName.replace("'", "''"))
                    .select('ViewQuery').get()
                    .then((view) => {
                        // Logger.writeJSON(view);
                        Logger.write('LiveSPClient.getViewQuery -> ' + view.ViewQuery);
                        this.viewQuery = view.ViewQuery;
                        resolve(this.viewQuery);
                    })
                    .catch((err) => {
                        Logger.write('Error reading view query:', LogLevel.Error);
                        Logger.writeJSON(err, LogLevel.Error);
                        resolve(undefined);
                    });
            } else {
                resolve(undefined);
            }
        });
    }
    public getFields(): Promise<ISPField[]> {
        return new Promise<ISPField[]>((resolve) => {
            if (this.fields) {
                resolve(this.fields);
            } else {
                // Logger.write('LiveSPClient.getFields() - reading fields');
                this.fields = [];
                if (this.list) {
                    this.list.fields
                        // TODO: figure out filter possibilities (fairly undocumented?)
                        .filter('Hidden eq false')
                        // and (FieldTypeKind eq 2 or FieldTypeKind eq 3 or (FieldTypeKind eq 17 and OutputType eq 2))
                        .select('*') // required to select all, as only then we get OutputType
                        // .select('Title', 'InternalName', 'FieldTypeKind', 'OutputType')
                        .get()
                        .then((spqueryresults) => {
                            // Logger.writeJSON(spqueryresults);
                            // strange, iterating the spqueryresults with foreach does not work?
                            for (let i: number = 0; i < spqueryresults.length; i++) {
                                const fieldinfo: IFieldInfo = spqueryresults[i];
                                if (
                                    (
                                        fieldinfo.FieldTypeKind === FieldTypes.Text ||
                                        fieldinfo.FieldTypeKind === FieldTypes.Note ||
                                        fieldinfo.FieldTypeKind === FieldTypes.Choice ||
                                        (
                                            fieldinfo.FieldTypeKind === FieldTypes.Calculated &&
                                            // tslint:disable-next-line:no-string-literal
                                            fieldinfo['OutputType'] === FieldTypes.Text
                                        )
                                    ) && (
                                        (fieldinfo.InternalName !== '_UIVersionString') &&
                                        (fieldinfo.InternalName !== 'Version')
                                    )
                                ) {
                                    // tslint:disable-next-line:max-line-length
                                    this.fields.push({ DisplayName: fieldinfo.Title, InternalName: fieldinfo.InternalName });
                                }
                            }
                            resolve(this.fields);
                        })
                        .catch((err) => {
                            Logger.write('Error reading fields:', LogLevel.Error);
                            Logger.writeJSON(err, LogLevel.Error);
                            resolve(this.fields);
                        });
                } else {
                    resolve(this.fields);
                }
            }
        });
    }
    public getContents(): Promise<string[]> {
        Logger.write('LiveSPClient.getContents()');
        // console.trace();
        return new Promise<string[]>((resolve) => {
            if (this.contents) {
                Logger.write('LiveSPClient.getContents() returns cached data');
                resolve(this.contents);
            } else {
                Logger.write('LiveSPClient.getContents() reading data');
                this.contents = [];
                if (!this.fieldInternalName) {
                    Logger.write('fieldInternalName not set');
                }
                if (this.fieldInternalName && this.list) {
                    this.getViewQuery()
                        .then((q: string) => {
                            const viewFields: string =
                                `<ViewFields><FieldRef Name='${this.fieldInternalName}'/></ViewFields>`;
                            const xml: string = `<View>${viewFields}<Query>${this.viewQuery}</Query></View>`;
                            // Logger.write('LiveSPClient.getContents() using ViewXml: ' + xml);
                            this.list
                                .getItemsByCAMLQuery({ 'ViewXml': xml })
                                // tslint:disable-next-line:no-any
                                .then((items: any) => {
                                    Logger.write('LiveSPClient.getContents() receiving data');
                                    Logger.writeJSON(items);
                                    // tslint:disable-next-line:no-any
                                    items.forEach((item: any) => {
                                        this.contents.push(item[this.fieldInternalName]);
                                    });
                                    resolve(this.contents);
                                });
                        });
                } else {
                    resolve(this.contents);
                }
            }
        });
    }

}