import { SPHttpClient } from '@microsoft/sp-http';
import { ILikeService, ILikeState } from './ILikeService';
import { getJson, post } from './restHelpers';

interface ILikedByInformation {
  isLikedByUser: boolean;
  likeCount: number;
  likedBy?: { name: string }[];
}

/**
 * Uses the modern page's built-in likes (the same data as the
 * "Like" button in the page footer and the page's social bar).
 */
export class PageLikeService implements ILikeService {
  private readonly _pageApi: string;

  constructor(private readonly _client: SPHttpClient, webAbsoluteUrl: string, pageItemId: number) {
    this._pageApi = `${webAbsoluteUrl}/_api/sitepages/pages(${pageItemId})`;
  }

  public async getState(): Promise<ILikeState> {
    const info = await getJson<ILikedByInformation>(
      this._client,
      `${this._pageApi}/likedByInformation?$expand=likedBy`
    );
    return {
      count: info.likeCount || 0,
      likedByMe: !!info.isLikedByUser,
      likers: (info.likedBy || []).map(l => l.name).filter(Boolean)
    };
  }

  public async like(): Promise<ILikeState> {
    await post(this._client, `${this._pageApi}/like`);
    return this.getState();
  }

  public async unlike(): Promise<ILikeState> {
    await post(this._client, `${this._pageApi}/unlike`);
    return this.getState();
  }
}
