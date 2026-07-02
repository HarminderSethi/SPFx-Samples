import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';
import { UI } from './constants/constants';
import { ShareDialog } from './dialog/ShareDialog';
import type { IShareItem } from './models/interfaces';
import { GraphService } from './services/GraphService';
import { getThemeColor } from '../../common/ThemeHelper';

export interface IShareViaTeamsChatCommandSetProperties {}

const LOG_SOURCE = 'ShareViaTeamsChatCommandSet';

const buildSharePointUrl = (webAbsoluteUrl: string, path: string): string => {
  if (!path) return webAbsoluteUrl;
  if (/^https?:\/\//i.test(path)) return path;

  try {
    const origin = new URL(webAbsoluteUrl).origin;
    const webPath = new URL(webAbsoluteUrl).pathname.replace(/\/+$/g, '').replace(/^\/+/, '');
    const decodedPath = decodeURIComponent(path);
    const normalizedPath = decodedPath.replace(/^\/+/, '');
    const combinedPath = normalizedPath.startsWith(webPath)
      ? normalizedPath
      : `${webPath}/${normalizedPath}`;

    return encodeURI(`${origin}/${combinedPath.replace(/\/+$/g, '/')}`);
  } catch {
    const fallbackUrl = path.startsWith('/') ? `${webAbsoluteUrl}${path}` : `${webAbsoluteUrl}/${path}`;
    return encodeURI(fallbackUrl);
  }
};

const getItemOpenUrl = (webAbsoluteUrl: string, listId: string | undefined, itemId: number): string => {
  const listQuery = listId ? `&ListId=${encodeURIComponent(listId)}` : '';
  return buildSharePointUrl(webAbsoluteUrl, `/_layouts/15/listform.aspx?PageType=4${listQuery}&ID=${itemId}`);
};

const getDocumentOpenUrl = (webAbsoluteUrl: string, serverRelativeUrl: string): string => {
  if (!serverRelativeUrl) return webAbsoluteUrl;
  const separator = serverRelativeUrl.includes('?') ? '&' : '?';
  return buildSharePointUrl(webAbsoluteUrl, `${serverRelativeUrl}${separator}web=1`);
};

const getRowValue = (row: { getValueByName(name: string): unknown }, name: string): string => {
  const value = row.getValueByName(name);
  return value === undefined || value === null ? '' : String(value);
};

const isLibraryRow = (row: { getValueByName(name: string): unknown }, listIsLibrary: boolean): boolean => {
  if (listIsLibrary) return true;
  const contentTypeId = getRowValue(row, 'ContentTypeId');
  return /^0x0101/i.test(contentTypeId) || /^0x0120/i.test(contentTypeId);
};

const getSendToTeamsIconUrl = (): string => {
  const themeColor = getThemeColor('themeDarkAlt');
  const fillColor = typeof themeColor === 'string' && themeColor.trim()
    ? themeColor.trim()
    : '#6264A7';

  const svgMarkup = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048">
    <path d="M1760 704q-47 0-87-17t-71-48-48-71-18-88q0-46 17-87t48-71 71-48 88-18q46 0 87 17t71 48 48 72 18 87q0 47-17 87t-48 71-72 48-87 18zm0-320q-40 0-68 28t-28 68q0 40 28 68t68 28q40 0 68-28t28-68q0-40-28-68t-68-28zm288 480v476q0 66-25 124t-68 102-102 69-125 25q-38 0-77-9t-73-28q-25 81-73 147t-112 114-143 74-162 26q-98 0-184-34t-154-94-112-142-58-178H85q-35 0-60-25t-25-60V597q0-35 25-60t60-25h733q-29-61-29-128 0-62 23-116t64-95 95-64 117-24q62 0 116 23t95 64 64 95 24 117q0 62-23 116t-64 95-95 64-117 24q-16 0-32-2t-32-5v92h928q40 0 68 28t28 68zm-960-651q-35 0-66 13t-55 37-36 55-14 66q0 35 13 66t37 55 54 36 67 14q35 0 66-13t54-37 37-54 14-67q0-35-13-66t-37-54-55-37-66-14zM592 848h192V688H240v160h192v512h160V848zm880 624V896h-448v555q0 35-25 60t-60 25H709q13 69 47 128t84 101 113 67 135 24q79 0 149-30t122-82 83-122 30-150zm448-132V896h-320v585q26 26 59 38t69 13q40 0 75-15t61-41 41-61 15-75z" fill="${fillColor}" />
  </svg>`;

  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgMarkup)}`;
};

export default class ShareViaTeamsChatCommandSet extends BaseListViewCommandSet<IShareViaTeamsChatCommandSetProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ShareViaTeamsChatCommandSet');

    const command: Command = this.tryGetCommand(UI.COMMAND_KEY);
    command.iconImageUrl = getSendToTeamsIconUrl();
    command.visible = false;

    return GraphService.initialize(this.context.msGraphClientFactory)
      .then(() => undefined)
      .catch((error: unknown) => {
        Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
      });
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const command: Command = this.tryGetCommand(UI.COMMAND_KEY);
    if (command) {
      command.visible = (event.selectedRows?.length ?? 0) >= 1;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== UI.COMMAND_KEY) return;
    if (!event.selectedRows || event.selectedRows.length === 0) return;

    this._showDialog(event).catch((error: unknown) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    });
  }

  private async _isCurrentListLibrary(): Promise<boolean> {
    const listId = this.context.pageContext.list?.id?.toString();
    if (!listId) return false;

    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')?$select=BaseType,BaseTemplate`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return false;
      const listInfo = (await response.json()) as { BaseType?: number; BaseTemplate?: number };
      return listInfo.BaseType === 1 || listInfo.BaseTemplate === 101;
    } catch {
      return false;
    }
  }

  private async _getItemModifiedDate(listId: string | undefined, itemId: number): Promise<string> {
    if (!listId) return '';

    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})?$select=Modified`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return '';
      const data = (await response.json()) as { Modified?: string };
      return data.Modified ?? '';
    } catch {
      return '';
    }
  }

  private async _showDialog(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    const listIsLibrary = await this._isCurrentListLibrary();

    const listTitle = this.context.pageContext.list?.title ?? '';
    const webAbsoluteUrl = this.context.pageContext.web.absoluteUrl;

    const listId = this.context.pageContext.list?.id?.toString();
    const items: IShareItem[] = await Promise.all(event.selectedRows.map(async (row) => {
      const id = Number(getRowValue(row, 'ID') || getRowValue(row, 'Id') || '0');
      const title = getRowValue(row, 'Title') || getRowValue(row, 'FileLeafRef') || `Item ${id}`;
      const fileRef = getRowValue(row, 'FileRef');
      const fileLeafRef = getRowValue(row, 'FileLeafRef');
      const isLibrary = isLibraryRow(row, listIsLibrary);
      const modifiedFromRow = [
        'Modified',
        'Last_x0020_Modified',
        'LastModified',
        'LastModifiedTime',
        'Modified_x0020_Date'
      ]
        .map((fieldName) => getRowValue(row, fieldName))
        .find((value) => value.length > 0) ?? '';
      const modified = modifiedFromRow || (isLibrary ? '' : await this._getItemModifiedDate(listId, id));
      const absoluteUrl = isLibrary
        ? getDocumentOpenUrl(webAbsoluteUrl, fileRef)
        : getItemOpenUrl(webAbsoluteUrl, listId, id);

      return {
        id,
        title,
        fileRef,
        fileLeafRef,
        listTitle,
        absoluteUrl,
        modified,
        isLibrary
      };
    }));

    const dialog = new ShareDialog({ items });
    dialog.show().catch((error: unknown) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    });
  }
}
