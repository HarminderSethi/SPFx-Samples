import * as React from 'react';
import type {
  IChannelResponse,
  IChatMember,
  IChatResponse,
  IDestination,
  ITeamResponse
} from '../models/interfaces';
import { TeamsService } from '../services/TeamsService';

interface IUseTeamsDestinationsResult {
  destinations: IDestination[];
  loading: boolean;
  error: string | undefined;
  chatsError: boolean;
}

const getMemberEmail = (member: IChatMember): string =>
  (member.email ?? member.userPrincipalName ?? '').trim();

const getMemberName = (member: IChatMember): string =>
  member.displayName?.trim() || getMemberEmail(member) || 'Unknown user';

const getUniqueChatMembers = (members: IChatMember[] = []): IChatMember[] => {
  const seen = new Set<string>();
  return members.filter((member) => {
    const email = getMemberEmail(member).toLowerCase();
    const key = email || member.userId || member.displayName;
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

const buildChatDisplayName = (chat: IChatResponse): string => {
  if (chat.topic && chat.topic.trim().length > 0) return chat.topic;
  const names = getUniqueChatMembers(chat.members).map(getMemberName).filter(Boolean);
  if (names.length === 0) return 'Untitled chat';
  return names.join(', ');
};

const getChatDedupeKey = (chat: IChatResponse): string => {
  const emails = getUniqueChatMembers(chat.members)
    .map((member) => getMemberEmail(member).toLowerCase())
    .filter(Boolean)
    .sort();

  return emails.length > 0 ? emails.join('|') : chat.id;
};

const mapChatsToDestinations = (chats: IChatResponse[]): IDestination[] => {
  const seenChats = new Set<string>();
  const destinations: IDestination[] = [];

  chats.forEach((chat) => {
    const chatKey = getChatDedupeKey(chat);
    if (seenChats.has(chatKey)) return;
    seenChats.add(chatKey);

    const uniqueMembers = getUniqueChatMembers(chat.members);
    const memberNames = uniqueMembers.map(getMemberName);
    const memberEmails = uniqueMembers.map(getMemberEmail).filter(Boolean);
    destinations.push({
      id: `chat:${chat.id}`,
      displayName: buildChatDisplayName(chat),
      type: 'chat',
      chatId: chat.id,
      memberNames,
      memberEmails
    });
  });

  return destinations;
};

const mapTeamsAndChannelsToDestinations = (
  teams: ITeamResponse[],
  channelsByTeam: Map<string, IChannelResponse[]>
): IDestination[] => {
  const result: IDestination[] = [];
  for (const team of teams) {
    const channels = channelsByTeam.get(team.id) ?? [];
    for (const channel of channels) {
      result.push({
        id: `channel:${team.id}:${channel.id}`,
        displayName: `${team.displayName} › ${channel.displayName}`,
        type: 'channel',
        teamId: team.id,
        channelId: channel.id
      });
    }
  }
  return result;
};

export function useTeamsDestinations(service: TeamsService): IUseTeamsDestinationsResult {
  const [destinations, setDestinations] = React.useState<IDestination[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [chatsError, setChatsError] = React.useState<boolean>(false);

  React.useEffect(() => {
    let cancelled = false;
    let teamsLoaded = false;
    let chatsLoaded = false;

    setLoading(true);
    setError(undefined);
    setChatsError(false);

    const updateLoadingState = (): void => {
      if (teamsLoaded && chatsLoaded) {
        setLoading(false);
      }
    };

    const withTimeout = <T>(promise: Promise<T>, ms: number, fallback: T): Promise<{ success: boolean; data: T }> => {
      let timeoutId: ReturnType<typeof setTimeout>;
      const timeoutPromise = new Promise<{ success: boolean; data: T }>((resolve) => {
        timeoutId = setTimeout(() => {
          resolve({ success: false, data: fallback });
        }, ms);
      });
      const normalPromise = promise
        .then((data) => ({ success: true, data }))
        .catch(() => ({ success: false, data: fallback }));
      return Promise.race([normalPromise, timeoutPromise]).then((res) => {
        clearTimeout(timeoutId);
        return res;
      });
    };

    // 1. Fetch Chats (with timeout)
    const loadChats = async (): Promise<void> => {
      try {
        const result = await withTimeout<IChatResponse[]>(service.getUserChats(), 8000, []);
        if (cancelled) return;

        if (!result.success) {
          setChatsError(true);
        }

        const chatDestinations = mapChatsToDestinations(result.data);
        setDestinations((prev) => {
          const withoutChats = prev.filter((d) => d.type !== 'chat');
          return [...chatDestinations, ...withoutChats];
        });
      } catch {
        if (cancelled) return;
        setChatsError(true);
      } finally {
        chatsLoaded = true;
        if (!cancelled) updateLoadingState();
      }
    };

    // 2. Fetch Teams and Channels
    const loadTeams = async (): Promise<void> => {
      try {
        const teams = await service.getJoinedTeams();
        if (cancelled) return;

        const channelLookups = await Promise.all(
          teams.map(async (team) => {
            try {
              const channels = await service.getTeamChannels(team.id);
              return [team.id, channels] as const;
            } catch {
              return [team.id, [] as IChannelResponse[]] as const;
            }
          })
        );

        if (cancelled) return;

        const channelsByTeam = new Map<string, IChannelResponse[]>(channelLookups);
        const teamDestinations = mapTeamsAndChannelsToDestinations(teams, channelsByTeam);

        setDestinations((prev) => {
          const withoutChannels = prev.filter((d) => d.type !== 'channel');
          return [...withoutChannels, ...teamDestinations];
        });
      } catch (err) {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : 'Failed to load teams.');
      } finally {
        teamsLoaded = true;
        if (!cancelled) updateLoadingState();
      }
    };

    // Run parallel loaders
    Promise.all([loadChats(), loadTeams()]).catch(() => {
      // Handled via local hooks states
    });

    return () => {
      cancelled = true;
    };
  }, [service]);

  return { destinations, loading, error, chatsError };
}
