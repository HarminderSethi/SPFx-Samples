import * as React from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Spinner,
  Text,
  Toast,
  ToastBody,
  ToastTitle
} from '@fluentui/react-components';
import { STRINGS, UI } from '../constants/constants';
import { useSendMessage } from '../hooks/useSendMessage';
import { useTeamsDestinations } from '../hooks/useTeamsDestinations';
import type {
  IDestination,
  ISharePanelProps,
  ISharePanelState
} from '../models/interfaces';
import { initialSharePanelState } from '../models/interfaces';
import { TeamsService } from '../services/TeamsService';
import { DestinationPicker } from './DestinationPicker';
import { MessageComposer } from './MessageComposer';
import { SelectedItemsList } from './SelectedItemsList';
import { useStyles } from './SharePanel.styles';

export const SharePanel: React.FC<ISharePanelProps> = ({ items, onDismiss, onShowToast }) => {
  const styles = useStyles();

  const teamsService = React.useRef<TeamsService>(new TeamsService()).current;

  const [shareItems, setShareItems] = React.useState(items);
  const [isDialogOpen, setIsDialogOpen] = React.useState(true);
  const [state, setState] = React.useState<ISharePanelState>(initialSharePanelState);
  const updateState = React.useCallback((partial: Partial<ISharePanelState>): void => {
    setState((prev) => ({ ...prev, ...partial }));
  }, []);

  const { destinations, loading: destinationsLoading, error: destinationsError, chatsError } =
    useTeamsDestinations(teamsService);

  React.useEffect(() => {
    updateState({ destinations, destinationsLoading, chatsError });
  }, [destinations, destinationsLoading, chatsError, updateState]);

  React.useEffect(() => {
    if (destinationsError) {
      updateState({ sendStatus: 'error', errorMessage: destinationsError });
    }
  }, [destinationsError, updateState]);

  const { status, error, send } = useSendMessage(teamsService);
  const previousStatusRef = React.useRef<"idle" | "sending" | "success" | "error">('idle');

  React.useEffect(() => {
    if (status === 'sending' || status === 'success' || status === 'error') {
      updateState({ sendStatus: status, errorMessage: error ?? '' });
    }
  }, [status, error, updateState]);

  React.useEffect(() => {
    if (status === previousStatusRef.current) return;

    if (status === 'success') {
      onShowToast?.(
        <Toast>
          <ToastTitle>{STRINGS.SUCCESS_MESSAGE}</ToastTitle>
         
        </Toast>,
        'success'
      );
      setIsDialogOpen(false);
      onDismiss();
    } else if (status === 'error') {
      const details = error ?? 'Unknown error';
      console.error('Share to Teams failed', details);
      onShowToast?.(
        <Toast>
          <ToastTitle>We could not send the message.</ToastTitle>
          <ToastBody>Please try again.</ToastBody>
        </Toast>,
        'error'
      );
    }

    previousStatusRef.current = status;
  }, [error, onDismiss, onShowToast, status]);

  const handleSearchChange = React.useCallback(
    (query: string): void => {
      updateState({ searchQuery: query });
    },
    [updateState]
  );

  const handleDestinationsSelected = React.useCallback(
    (destinations: IDestination[]): void => {
      updateState({ selectedDestinations: destinations });
    },
    [updateState]
  );

  const handleMessageChange = React.useCallback(
    (value: string): void => {
      updateState({ messageText: value });
    },
    [updateState]
  );

  const handleRemoveItem = React.useCallback((itemId: number): void => {
    setShareItems((prev) => prev.filter((item) => item.id !== itemId));
  }, []);

  const handleSend = React.useCallback((): void => {
    if (state.selectedDestinations.length === 0 || shareItems.length === 0) return;
    send(state.selectedDestinations, shareItems, state.messageText).catch(() => {
      // Errors are surfaced via send() state into status/error.
    });
  }, [send, state.selectedDestinations, state.messageText, shareItems]);

  const tooManyItems = shareItems.length > UI.MAX_SHARE_ITEMS;
  const isSuccess = state.sendStatus === 'success';
  const footerActionLabel = isSuccess ? STRINGS.CLOSE_BUTTON : STRINGS.CANCEL_BUTTON;
  const sendDisabled =
    shareItems.length === 0 ||
    tooManyItems ||
    state.selectedDestinations.length === 0 ||
    state.sendStatus === 'sending' ||
    state.messageText.length > UI.MESSAGE_MAX_LENGTH;

  return (
    <Dialog open={isDialogOpen} modalType="modal" onOpenChange={(_event, data) => !data.open && onDismiss()}>
      <DialogSurface className={styles.surface}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                aria-label={STRINGS.CLOSE_BUTTON}
                onClick={onDismiss}
              >
                ✕
              </Button>
            }
          >
            {STRINGS.PANEL_TITLE}
          </DialogTitle>
          <DialogContent>
            <div className={styles.body}>
              <SelectedItemsList items={shareItems} onRemoveItem={handleRemoveItem} />
              <DestinationPicker
                destinations={state.destinations}
                selectedDestinations={state.selectedDestinations}
                loading={state.destinationsLoading}
                searchQuery={state.searchQuery}
                onSearchChange={handleSearchChange}
                onDestinationsSelected={handleDestinationsSelected}
                chatsError={state.chatsError}
              />
              <MessageComposer
                value={state.messageText}
                onChange={handleMessageChange}
                maxLength={UI.MESSAGE_MAX_LENGTH}
              />
              {state.sendStatus === 'sending' && (
                <div className={styles.sendingState}>
                  <Spinner size="tiny" />
                  <Text>{STRINGS.SENDING_LABEL}</Text>
                </div>
              )}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onDismiss}>
              {footerActionLabel}
            </Button>
            <Button appearance="primary" disabled={sendDisabled} onClick={handleSend}>
              {STRINGS.SEND_BUTTON}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

SharePanel.displayName = 'SharePanel';
