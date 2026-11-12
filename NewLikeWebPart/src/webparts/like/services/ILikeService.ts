export interface ILikeState {
  /** Total number of likes. */
  count: number;
  /** Whether the current user has liked. */
  likedByMe: boolean;
  /** Display names of people who liked (may be a subset). */
  likers: string[];
}

export interface ILikeService {
  getState(): Promise<ILikeState>;
  like(): Promise<ILikeState>;
  unlike(): Promise<ILikeState>;
}

/** Error raised when the REST call fails, carrying the HTTP status. */
export class LikeServiceError extends Error {
  public readonly status: number;

  constructor(message: string, status: number) {
    super(message);
    this.status = status;
    // Keep instanceof working when targeting ES5.
    Object.setPrototypeOf(this, LikeServiceError.prototype);
  }
}
