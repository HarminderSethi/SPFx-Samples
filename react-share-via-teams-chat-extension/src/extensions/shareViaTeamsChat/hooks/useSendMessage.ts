import * as React from 'react';
import type { IDestination, IShareItem, SendStatus } from '../models/interfaces';
import { AdaptiveCardBuilder } from '../services/AdaptiveCardBuilder';
import { TeamsService } from '../services/TeamsService';

interface IUseSendMessageResult {
  status: SendStatus;
  error: string | undefined;
  send: (destinations: IDestination[], items: IShareItem[], message: string) => Promise<void>;
  reset: () => void;
}

export function useSendMessage(
  service: TeamsService
): IUseSendMessageResult {
  const [status, setStatus] = React.useState<SendStatus>('idle');
  const [error, setError] = React.useState<string | undefined>(undefined);

  const send = React.useCallback(
    async (destinations: IDestination[], items: IShareItem[], message: string): Promise<void> => {
      setStatus('sending');
      setError(undefined);
      try {
        if (destinations.length === 0) {
          throw new Error('Select at least one destination.');
        }
        const body = AdaptiveCardBuilder.buildMessageBody(items, message);
        await Promise.all(destinations.map(async (destination) => {
          if (destination.type === 'chat' && destination.chatId) {
            await service.sendChatMessage(destination.chatId, body);
            return;
          }
          if (
            destination.type === 'channel' &&
            destination.teamId &&
            destination.channelId
          ) {
            await service.sendChannelMessage(destination.teamId, destination.channelId, body);
            return;
          }
          throw new Error('Selected destination is incomplete.');
        }));
        setStatus('success');
      } catch (err) {
        setStatus('error');
        setError(err instanceof Error ? err.message : 'Failed to send message.');
      }
    },
    [service]
  );

  const reset = React.useCallback((): void => {
    setStatus('idle');
    setError(undefined);
  }, []);

  return { status, error, send, reset };
}
