import { GRAPH_ENDPOINTS } from '../constants/constants';
import type {
  IChannelResponse,
  IChatMessageBody,
  IChatResponse,
  IGraphCollectionResponse,
  ITeamResponse
} from '../models/interfaces';
import { GraphService } from './GraphService';

/**
 * Teams-specific façade over `GraphService`. Caches chat and team
 * lookups for the lifecycle of the consumer (typically the open dialog).
 */
export class TeamsService {
  private _chatsCache: IChatResponse[] | undefined;
  private _teamsCache: ITeamResponse[] | undefined;
  private readonly _channelsCache = new Map<string, IChannelResponse[]>();

  private async _getAllPages<T>(endpoint: string): Promise<T[]> {
    const results: T[] = [];
    let nextEndpoint: string | undefined = endpoint;

    while (nextEndpoint) {
      const response: IGraphCollectionResponse<T> = await GraphService.getInstance().get<IGraphCollectionResponse<T>>(nextEndpoint);
      results.push(...(response.value ?? []));
      nextEndpoint = response['@odata.nextLink'];
    }

    return results;
  }

  public async getUserChats(forceRefresh = false): Promise<IChatResponse[]> {
    if (!forceRefresh && this._chatsCache) return this._chatsCache;
    this._chatsCache = await this._getAllPages<IChatResponse>(GRAPH_ENDPOINTS.ME_CHATS);
    return this._chatsCache;
  }

  public async getJoinedTeams(forceRefresh = false): Promise<ITeamResponse[]> {
    if (!forceRefresh && this._teamsCache) return this._teamsCache;
    this._teamsCache = await this._getAllPages<ITeamResponse>(GRAPH_ENDPOINTS.ME_JOINED_TEAMS);
    return this._teamsCache;
  }

  public async getTeamChannels(teamId: string, forceRefresh = false): Promise<IChannelResponse[]> {
    const cached = this._channelsCache.get(teamId);
    if (!forceRefresh && cached) return cached;
    const channels = await this._getAllPages<IChannelResponse>(GRAPH_ENDPOINTS.TEAM_CHANNELS(teamId));
    this._channelsCache.set(teamId, channels);
    return channels;
  }

  public async sendChatMessage(chatId: string, body: IChatMessageBody): Promise<void> {
    await GraphService.getInstance().post<unknown>(GRAPH_ENDPOINTS.CHAT_MESSAGES(chatId), body);
  }

  public async sendChannelMessage(
    teamId: string,
    channelId: string,
    body: IChatMessageBody
  ): Promise<void> {
    await GraphService.getInstance().post<unknown>(
      GRAPH_ENDPOINTS.CHANNEL_MESSAGES(teamId, channelId),
      body
    );
  }

  public clearCache(): void {
    this._chatsCache = undefined;
    this._teamsCache = undefined;
    this._channelsCache.clear();
  }
}
