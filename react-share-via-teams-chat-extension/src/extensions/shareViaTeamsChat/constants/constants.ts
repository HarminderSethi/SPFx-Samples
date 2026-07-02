export const GRAPH_ENDPOINTS = {
  ME_CHATS: '/me/chats?$expand=members&$top=50',
  ME_JOINED_TEAMS: '/me/joinedTeams',
  TEAM_CHANNELS: (teamId: string): string => `/teams/${teamId}/channels`,
  CHAT_MESSAGES: (chatId: string): string => `/chats/${chatId}/messages`,
  CHANNEL_MESSAGES: (teamId: string, channelId: string): string =>
    `/teams/${teamId}/channels/${channelId}/messages`
} as const;

export const UI = {
  DEBOUNCE_MS: 300,
  MESSAGE_MAX_LENGTH: 280,
  MAX_SHARE_ITEMS: 10,
  COMMAND_KEY: 'SEND_TO_TEAMS'
} as const;

export const ADAPTIVE_CARD = {
  SCHEMA: 'http://adaptivecards.io/schemas/adaptive-card.json',
  VERSION: '1.5',
  CONTENT_TYPE: 'application/vnd.microsoft.card.adaptive' as const
} as const;

export const STRINGS = {
  COMMAND_TITLE: 'Send to Teams',
  PANEL_TITLE: 'Share via Teams Chat',
  SEARCH_PLACEHOLDER: 'Search chats, teams, or channels…',
  MESSAGE_PLACEHOLDER: 'Add an optional message…',
  SEND_BUTTON: 'Send',
  CANCEL_BUTTON: 'Cancel',
  CLOSE_BUTTON: 'Close',
  SENDING_LABEL: 'Sending to Teams…',
  SUCCESS_MESSAGE: 'Message sent successfully!',
  NO_RESULTS: 'No matching chats, teams, or channels found.',
  CHAT_PERMISSION_HINT: 'Chats will appear after Microsoft Graph chat consent is granted.',
  ITEM_CARD_HEADING: "You're sharing",
  NO_ITEMS_SELECTED: 'No items selected.',
  RECENT_CHATS_GROUP: 'Recent Chats',
  TEAMS_CHANNELS_GROUP: 'Teams & Channels',
  CLEAR_DESTINATIONS: 'Clear all',
  ITEMS_SELECTED: (count: number): string =>
    `${count} ${count === 1 ? 'item' : 'items'} selected`,
  TOO_MANY_ITEMS: (count: number, max: number): string =>
    `Select ${max} or fewer items to send to Teams. ${count} items are currently selected.`,
  CHAR_COUNTER: (current: number, max: number): string => `${current}/${max} chars`
} as const;
