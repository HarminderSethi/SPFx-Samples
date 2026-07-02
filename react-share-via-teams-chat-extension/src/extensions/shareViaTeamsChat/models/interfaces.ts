// ─── Selected SharePoint Item ────────────────────────────
export interface IShareItem {
  id: number;
  title: string;
  fileRef: string;
  fileLeafRef: string;
  listTitle: string;
  absoluteUrl: string;
  modified: string;
  isLibrary?: boolean;
}

// ─── Destination (unified) ───────────────────────────────
export type DestinationType = 'chat' | 'channel';

export interface IDestination {
  id: string;
  displayName: string;
  type: DestinationType;
  teamId?: string;
  channelId?: string;
  chatId?: string;
  memberNames?: string[];
  memberEmails?: string[];
}

// ─── Teams API Models ────────────────────────────────────
export interface IChatMember {
  displayName: string;
  userId: string;
  email?: string;
  userPrincipalName?: string;
}

export interface IChatResponse {
  id: string;
  topic: string | undefined;
  chatType: 'oneOnOne' | 'group' | 'meeting';
  members: IChatMember[];
}

export interface ITeamResponse {
  id: string;
  displayName: string;
  description: string;
}

export interface IChannelResponse {
  id: string;
  displayName: string;
  membershipType: string;
}

export interface IGraphCollectionResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

// ─── Adaptive Card ───────────────────────────────────────
export interface IAdaptiveCardAttachment {
  id: string;
  contentType: 'application/vnd.microsoft.card.adaptive';
  contentUrl: undefined;
  content: string;
}

export interface IChatMessageBody {
  body: {
    contentType: 'html' | 'text';
    content: string;
  };
  attachments?: IAdaptiveCardAttachment[];
}

// ─── Consolidated Component State ────────────────────────
export type SendStatus = 'idle' | 'sending' | 'success' | 'error';

export interface ISharePanelState {
  selectedDestinations: IDestination[];
  messageText: string;
  sendStatus: SendStatus;
  errorMessage: string;
  destinations: IDestination[];
  destinationsLoading: boolean;
  searchQuery: string;
  chatsError: boolean;
}

export const initialSharePanelState: ISharePanelState = {
  selectedDestinations: [],
  messageText: '',
  sendStatus: 'idle',
  errorMessage: '',
  destinations: [],
  destinationsLoading: true,
  searchQuery: '',
  chatsError: false
};

// ─── Component Props ─────────────────────────────────────
export interface ISharePanelProps {
  items: IShareItem[];
  onDismiss: () => void;
  onShowToast?: (content: React.ReactNode, intent?: 'success' | 'error' | 'info' | 'warning') => void;
}

export interface ISelectedItemsListProps {
  items: IShareItem[];
  onRemoveItem?: (itemId: number) => void;
}

export interface IDestinationPickerProps {
  destinations: IDestination[];
  selectedDestinations: IDestination[];
  loading: boolean;
  searchQuery: string;
  onSearchChange: (query: string) => void;
  onDestinationsSelected: (destinations: IDestination[]) => void;
  chatsError: boolean;
}

export interface IMessageComposerProps {
  value: string;
  onChange: (value: string) => void;
  maxLength: number;
}

// ─── Dialog Props ────────────────────────────────────────
export interface IShareDialogProps {
  items: IShareItem[];
}
