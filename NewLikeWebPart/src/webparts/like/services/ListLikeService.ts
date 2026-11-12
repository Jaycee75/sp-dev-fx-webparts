import { SPHttpClient } from '@microsoft/sp-http';
import { ILikeService, ILikeState, LikeServiceError } from './ILikeService';
import { getJson, odataString, post } from './restHelpers';

interface ILikeItem {
  Id: number;
  AuthorId: number;
  Author?: { Title: string };
}

/**
 * Stores likes in a SharePoint custom list: one item per user per like key.
 * Title = like key, Created By = the user who liked.
 */
export class ListLikeService implements ILikeService {
  private readonly _listApi: string;
  private _listEnsured = false;

  constructor(
    private readonly _client: SPHttpClient,
    private readonly _webUrl: string,
    private readonly _listTitle: string,
    private readonly _likeKey: string,
    private readonly _currentUserId: number
  ) {
    this._listApi = `${_webUrl}/_api/web/lists/getbytitle('${odataString(_listTitle)}')`;
  }

  public async getState(): Promise<ILikeState> {
    const items = await this._getItems();
    return this._toState(items);
  }

  public async like(): Promise<ILikeState> {
    await this._ensureList();
    const items = await this._getItems();
    if (!items.some(i => i.AuthorId === this._currentUserId)) {
      await post(this._client, `${this._listApi}/items`, { Title: this._likeKey });
    }
    return this.getState();
  }

  public async unlike(): Promise<ILikeState> {
    const mine = (await this._getItems()).filter(i => i.AuthorId === this._currentUserId);
    for (const item of mine) {
      await post(this._client, `${this._listApi}/items(${item.Id})`, undefined, {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      });
    }
    return this.getState();
  }

  private async _getItems(): Promise<ILikeItem[]> {
    const url =
      `${this._listApi}/items?$select=Id,AuthorId,Author/Title&$expand=Author` +
      `&$filter=Title eq '${odataString(this._likeKey)}'&$orderby=Created desc&$top=5000`;
    try {
      const result = await getJson<{ value: ILikeItem[] }>(this._client, url);
      this._listEnsured = true;
      return result.value || [];
    } catch (e) {
      // The list doesn't exist yet: nobody has liked anything.
      if (e instanceof LikeServiceError && e.status === 404) {
        return [];
      }
      throw e;
    }
  }

  private async _ensureList(): Promise<void> {
    if (this._listEnsured) {
      return;
    }
    try {
      await getJson(this._client, `${this._listApi}?$select=Id`);
    } catch (e) {
      if (!(e instanceof LikeServiceError) || e.status !== 404) {
        throw e;
      }
      await post(this._client, `${this._webUrl}/_api/web/lists`, {
        Title: this._listTitle,
        Description: 'Likes recorded by the Like button web part. Title = like key.',
        BaseTemplate: 100,
        AllowContentTypes: false,
        ContentTypesEnabled: false
      });
    }
    this._listEnsured = true;
  }

  private _toState(items: ILikeItem[]): ILikeState {
    // One like per user, even if duplicates slipped in through concurrent clicks.
    const byUser = new Map<number, string>();
    items.forEach(i => {
      if (!byUser.has(i.AuthorId)) {
        byUser.set(i.AuthorId, i.Author?.Title || '');
      }
    });
    const likers: string[] = [];
    byUser.forEach(name => {
      if (name) {
        likers.push(name);
      }
    });
    return {
      count: byUser.size,
      likedByMe: byUser.has(this._currentUserId),
      likers
    };
  }
}
