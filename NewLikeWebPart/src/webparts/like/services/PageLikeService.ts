import { SPHttpClient } from '@microsoft/sp-http';
import { ILikeService, ILikeState, LikeServiceError } from './ILikeService';
import { getJson, post } from './restHelpers';

interface ILikedByInformation {
  isLikedByUser: boolean;
  likeCount: number;
  likedBy?: { name: string }[];
}

/**
 * Uses the modern page's built-in likes (the same data as the
 * "Like" button in the page footer and the page's social bar).
 *
 * The supported endpoints hang off the page's list item in Site Pages:
 *   GET  /_api/web/lists(guid'{list}')/GetItemById({id})/likedByInformation?$expand=likedby
 *   POST /_api/web/lists(guid'{list}')/GetItemById({id})/like  |  /unlike
 * The /_api/sitepages/pages({id}) variants are tried as a fallback.
 */
export class PageLikeService implements ILikeService {
  private readonly _apis: string[];

  constructor(private readonly _client: SPHttpClient, webAbsoluteUrl: string, listId: string, pageItemId: number) {
    const guid = listId.replace(/[{}]/g, '');
    this._apis = [
      `${webAbsoluteUrl}/_api/web/lists(guid'${guid}')/GetItemById(${pageItemId})`,
      `${webAbsoluteUrl}/_api/sitepages/pages(${pageItemId})`
    ];
  }

  public async getState(): Promise<ILikeState> {
    const info = await this._withFallback(async api => {
      try {
        return await getJson<ILikedByInformation>(this._client, `${api}/likedByInformation?$expand=likedby`);
      } catch (e) {
        // Some tenants reject the expand; the count alone is still useful.
        if (e instanceof LikeServiceError && e.status === 400) {
          return getJson<ILikedByInformation>(this._client, `${api}/likedByInformation`);
        }
        throw e;
      }
    });
    return {
      count: info.likeCount || 0,
      likedByMe: !!info.isLikedByUser,
      likers: (info.likedBy || []).map(l => l.name).filter(Boolean)
    };
  }

  public async like(): Promise<ILikeState> {
    await this._withFallback(api => post(this._client, `${api}/like`));
    return this.getState();
  }

  public async unlike(): Promise<ILikeState> {
    await this._withFallback(api => post(this._client, `${api}/unlike`));
    return this.getState();
  }

  /** Try each endpoint in turn, moving on only when one doesn't exist (404). */
  private async _withFallback<T>(call: (api: string) => Promise<T>): Promise<T> {
    let lastError: unknown;
    for (const api of this._apis) {
      try {
        return await call(api);
      } catch (e) {
        lastError = e;
        if (!(e instanceof LikeServiceError) || e.status !== 404) {
          throw e;
        }
      }
    }
    throw lastError;
  }
}
