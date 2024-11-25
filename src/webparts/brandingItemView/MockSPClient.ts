import { Logger } from '@pnp/logging';
import { Web } from '@pnp/sp';
import { ISPList, ISPView, ISPField, IDataSource } from './BrandingItemViewWebPart';

export default class MockSPClient implements IDataSource {
  private static _lists: ISPList[] = [
    { Title: 'Mock List', Id: '1' },
    { Title: 'Another List', Id: '2' },
    { Title: 'Last List', Id: '3' }
  ];
  private static _views: ISPView[] = [
    { Title: 'All Items', Id: '1' },
    { Title: 'Filtered Items', Id: '2' },
    { Title: 'New Items', Id: '3' }
  ];
  private static _fields: ISPField[] = [
    { DisplayName: 'Title', InternalName: 'Title' },
    { DisplayName: 'Nicely Formatted Text', InternalName: 'Text' }
  ];
  private static _contents: string[] = [
    'This is the first text element',
    'Here is some more text with <strong>strong</strong> or <em>emphasis</em> markup. Should work.',
    'And some last element. Nothing special'
  ];

  public setWeb(web: Web): void {
    // not implemented in Mockup
  }
  public setList(listName: string): void {
    // not implemented in Mockup
  }
  public setView(viewName: string): void {
    // not implemented in Mockup
  }
  public setField(internalName: string): void {
    // not implemented in Mockup
  }
  public getLists(): Promise<ISPList[]> {
    Logger.write('MockSPClient.getLists()');
    return new Promise<ISPList[]>((resolve) => {
      resolve(MockSPClient._lists);
    });
  }
  public getViews(): Promise<ISPView[]> {
    Logger.write('MockSPClient.getViews()');
    return new Promise<ISPView[]>((resolve) => {
      resolve(MockSPClient._views);
    });
  }
  public getFields(): Promise<ISPField[]> {
    Logger.write('MockSPClient.getFields()');
    return new Promise<ISPField[]>((resolve) => {
      resolve(MockSPClient._fields);
    });
  }
  public getContents(): Promise<string[]> {
    Logger.write('MockSPClient.getContents()');
    return new Promise<string[]>((resolve) => {
      resolve(MockSPClient._contents);
    });
  }
}