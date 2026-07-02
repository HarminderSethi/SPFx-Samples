# Share via Teams Chat — SPFx ListView Command Set Extension

## Implementation Plan

> A ListView Command Set extension for SharePoint lists and document libraries that surfaces a **"Send to Teams"** button in the toolbar. When invoked, it opens a modal dialog allowing users to search for a Teams channel, 1:1, or group chat, compose a quick message, and dispatch a neatly formatted Adaptive Card linking back to one or more selected list items. Supports **multi-select** — users can share multiple items at once.

---

## Technology Stack

| Layer | Technology | Version |
|---|---|---|
| Framework | SharePoint Framework (SPFx) | **1.23** (Heft toolchain) |
| UI Library | Fluent UI React v9 (`@fluentui/react-components`) | **^9.x** |
| Icons | `@fluentui/react-icons` | **^2.x** |
| React | React 17 + `react-dom` | **17.0.1** |
| API Client | `MSGraphClientV3` (built-in SPFx) | — |
| Language | TypeScript (strict mode) | **^4.x** |

> [!IMPORTANT]
> SPFx 1.23 ships with **React 17.0.1**. Do **not** use React 18 APIs (`createRoot`, concurrent features). All rendering must use `ReactDOM.render()`.

---

## Architecture Overview

```
src/
└── extensions/
    └── shareViaTeamsChat/
        │
        │  ── Entry Point ──────────────────────────────────────
        ├── ShareViaTeamsChatCommandSet.ts          # ListView Command Set (onInit / onExecute)
        │
        │  ── Dialog Bridge ────────────────────────────────────
        ├── dialog/
        │   └── ShareDialog.ts                      # BaseDialog subclass → ReactDOM.render (React 17)
        │
        │  ── React Components ─────────────────────────────────
        ├── components/
        │   ├── SharePanel.tsx                       # Root modal orchestrator (Dialog)
        │   ├── SharePanel.styles.ts                  # Scoped styles (makeStyles)
        │   ├── SelectedItemsList/
        │   │   ├── SelectedItemsList.tsx             # Multi-item preview list
        │   │   └── SelectedItemsList.styles.ts       # Scoped styles (makeStyles)
        │   ├── DestinationPicker/
        │   │   ├── DestinationPicker.tsx             # Combobox to search chats & channels
        │   │   └── DestinationPicker.styles.ts       # Scoped styles (makeStyles)
        │   ├── MessageComposer/
        │   │   ├── MessageComposer.tsx              # Textarea + send button
        │   │   └── MessageComposer.styles.ts
        │   └── StatusIndicator/
        │       ├── StatusIndicator.tsx               # Inline status (sending, sent, error)
        │       └── StatusIndicator.styles.ts
        │
        │  ── Services ─────────────────────────────────────────
        ├── services/
        │   ├── GraphService.ts                     # MSGraphClientV3 wrapper (singleton)
        │   ├── TeamsService.ts                     # Chats, channels, send message
        │   └── AdaptiveCardBuilder.ts              # Builds the Adaptive Card JSON payload
        │
        │  ── Hooks ────────────────────────────────────────────
        ├── hooks/
        │   ├── useTeamsChats.ts                    # Fetches & caches user's chats
        │   ├── useJoinedTeams.ts                   # Fetches user's joined teams & channels
        │   ├── useSendMessage.ts                   # Mutation hook for posting the message
        │   └── useDebounce.ts                      # Debounce utility for search input
        │
        │  ── Shared ───────────────────────────────────────────
        ├── models/
        │   └── interfaces.ts                       # All TypeScript interfaces & types
        ├── constants/
        │   └── constants.ts                        # Strings, Graph endpoints, icon keys
        └── loc/
            ├── en-us.js                            # Localised strings (English)
            └── myStrings.d.ts                      # String type declarations
```

---

## Microsoft Graph API Permissions

Declared in `config/package-solution.json` under `webApiPermissionRequests`. All permissions are **delegated** (signed-in user context).

| Permission | Scope | Purpose |
|---|---|---|
| `Chat.ReadWrite` | Delegated | List the user's 1:1 and group chats, send messages to chats |
| `Team.ReadBasic.All` | Delegated | List the user's joined Teams |
| `Channel.ReadBasic.All` | Delegated | List channels within a Team |
| `ChannelMessage.Send` | Delegated | Post messages (including Adaptive Cards) to a channel |
| `User.Read` | Delegated | Read signed-in user's profile (sender context) |

> [!IMPORTANT]
> After deployment, a SharePoint tenant admin must approve these permissions via **SharePoint Admin Center → API Access**.

---

## Detailed Component Breakdown

### 1. Command Set Entry Point

**File:** `ShareViaTeamsChatCommandSet.ts`

- Extends `BaseListViewCommandSet<{}>`.
- Registers a single command **`SEND_TO_TEAMS`** in the manifest.
- `onInit()` → initialises `GraphService` singleton via `this.context.msGraphClientFactory`.
- `onListViewUpdated()` → shows/hides the command based on `event.selectedRows.length >= 1` (**multi-select enabled**).
- `onExecute()` → iterates over `event.selectedRows` and maps each row to an `IShareItem`. Passes the full `IShareItem[]` array to the `ShareDialog`.

**Manifest configuration (`ShareViaTeamsChatCommandSet.manifest.json`):**
```json
{
  "items": {
    "SEND_TO_TEAMS": {
      "title": { "default": "Send to Teams" },
      "iconImageUrl": "data:image/svg+xml,...",
      "type": "command"
    }
  }
}
```

---

### 2. Dialog Bridge

**File:** `dialog/ShareDialog.ts`

- Extends `BaseDialog` from `@microsoft/sp-dialog`.
- Accepts `IShareDialogProps` containing the selected **items array** (`IShareItem[]`) and the `MSGraphClientV3` instance.
- `render()` → uses `ReactDOM.render()` (React 17 pattern — **not** `createRoot`).
- Wraps `<SharePanel />` inside `<FluentProvider theme={webLightTheme}>`.
- `onAfterClose()` → calls `ReactDOM.unmountComponentAtNode(this.domElement)` to prevent memory leaks.
- Passes a `close()` callback to React so the modal can dismiss itself.

---

### 3. React Components

#### 3a. SharePanel (Root Orchestrator)

**File:** `components/SharePanel.tsx`

| Responsibility | Detail |
|---|---|
| Layout | Uses Fluent UI `Dialog` (modal, `size: large`) with `DialogSurface`, `DialogBody`, `DialogTitle`, `DialogContent`, `DialogActions` |
| Sections | `DialogTitle` → title + close button; `DialogContent` → SelectedItemsList, DestinationPicker, MessageComposer, StatusIndicator; `DialogActions` → Send / Cancel buttons |
| State | **Single consolidated state object** via `ISharePanelState` (see below) |
| Flow | 1. User sees selected items list → 2. Picks destination → 3. Types optional message → 4. Sends |

**Consolidated state pattern:**
```typescript
interface ISharePanelState {
  selectedDestination: IDestination | null;
  messageText: string;
  sendStatus: SendStatus;
  errorMessage: string;
  destinations: IDestination[];         // merged chats + channels
  destinationsLoading: boolean;
  searchQuery: string;
  chatsError: boolean;
}

// Inside the component:
const [state, setState] = React.useState<ISharePanelState>(initialSharePanelState);

// Partial updates:
setState(prev => ({ ...prev, sendStatus: 'sending' }));
```

> [!IMPORTANT]
> All state in `SharePanel` is held in a **single state object** and updated via spread-merge. Child components receive values and callbacks as props — they do not own independent state.

---

#### 3b. SelectedItemsList (NEW)

**File:** `components/SelectedItemsList/SelectedItemsList.tsx`

- Receives `items: IShareItem[]` as props.
- Renders a compact, scrollable list of all selected items (supports **multi-select**).
- Each row shows: file-type icon, **item title**, list name, modified date.
- If a single item is selected, renders as a single `Card`. If multiple, renders as a vertical stack with a count badge (e.g., "3 items selected").
- Purely presentational — no internal state.

---

#### 3c. DestinationPicker

**File:** `components/DestinationPicker/DestinationPicker.tsx`

- Uses Fluent UI `Combobox` with `freeform` search.
- Receives `destinations`, `loading`, `searchQuery`, and `onSelect` / `onSearchChange` callbacks as props from `SharePanel`.
- Groups results under `OptionGroup` headers: **"Recent Chats"** / **"Teams & Channels"**.
- Implements `useDebounce(300ms)` on the search input for client-side filtering.
- Displays `Avatar` for chat participants and team icons for channels.

> [!NOTE]
> If a custom People Picker component is provided before development, it will be integrated into the destination picker flow as needed. The core picker here is for **Teams channels, chats, and groups** — not people.

---

#### 3d. MessageComposer

**File:** `components/MessageComposer/MessageComposer.tsx`

- Fluent UI `Textarea` bound to parent's `messageText` via props.
- Character counter with a soft limit (configurable via `constants.ts`, default: **280 chars**).
- Placeholder text providing contextual guidance.
- Lightweight component — receives `value` and `onChange` as props (controlled input). No internal state.

---

#### 3e. StatusIndicator

**File:** `components/StatusIndicator/StatusIndicator.tsx`

- Renders contextual feedback:
  - `idle` → nothing rendered.
  - `sending` → `Spinner` with "Sending to Teams…" label.
  - `success` → `MessageBar` (success) with "Message sent successfully" + auto-dismiss after 3s.
  - `error` → `MessageBar` (error) with retry affordance.
- Uses `SendStatus` type from `interfaces.ts`.
- Receives all values as props — no internal state.

---

### 4. Services Layer

#### 4a. GraphService

**File:** `services/GraphService.ts`

```typescript
export class GraphService {
  private static _instance: GraphService;
  private _client: MSGraphClientV3;

  public static async initialize(factory: MSGraphClientFactory): Promise<GraphService>;
  public static getInstance(): GraphService;
  public async get<T>(endpoint: string): Promise<T>;
  public async post<T>(endpoint: string, body: unknown): Promise<T>;
}
```

- **Singleton** — initialized once during `onInit()`, reused everywhere.
- Wraps all Graph calls with consistent error handling (`try/catch` + typed error responses).
- Uses `v1.0` endpoints only (no `/beta`).

---

#### 4b. TeamsService

**File:** `services/TeamsService.ts`

| Method | Graph Endpoint | Returns |
|---|---|---|
| `getUserChats()` | `GET /me/chats?$expand=members&$top=50` | `IChatDestination[]` |
| `getJoinedTeams()` | `GET /me/joinedTeams` | `ITeam[]` |
| `getTeamChannels(teamId)` | `GET /teams/{teamId}/channels` | `IChannel[]` |
| `sendChatMessage(chatId, body)` | `POST /chats/{chatId}/messages` | `void` |
| `sendChannelMessage(teamId, channelId, body)` | `POST /teams/{teamId}/channels/{channelId}/messages` | `void` |

- Consumes `GraphService.getInstance()`.
- All methods are `async` and return typed results.
- Caches `getUserChats()` and `getJoinedTeams()` results in-memory for the lifecycle of the dialog (avoids redundant calls on every keystroke).

---

#### 4c. AdaptiveCardBuilder

**File:** `services/AdaptiveCardBuilder.ts`

Generates a clean, branded Adaptive Card (schema `1.4`) containing:

```json
{
  "type": "AdaptiveCard",
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "version": "1.4",
  "body": [
    {
      "type": "ColumnSet",
      "columns": [
        {
          "type": "Column",
          "width": "auto",
          "items": [
            { "type": "Image", "url": "https://upload.wikimedia.org/wikipedia/commons/e/ec/Microsoft_Office_SharePoint_%282019%E2%80%932025%29.svg", "size": "Small", "width": "32px", "height": "32px" }
          ]
        },
        {
          "type": "Column",
          "width": "stretch",
          "items": [
            { "type": "TextBlock", "text": "Shared from SharePoint", "weight": "Bolder", "size": "Medium" },
            { "type": "TextBlock", "text": "<List Name>", "isSubtle": true, "spacing": "None" }
          ]
        }
      ]
    },
    { "type": "TextBlock", "text": "<Item Title>", "weight": "Bolder", "size": "Large", "wrap": true },
    { "type": "TextBlock", "text": "<User Message>", "wrap": true },
    {
      "type": "FactSet",
      "facts": [
        { "title": "Modified", "value": "<date>" },
        { "title": "Shared by", "value": "<current user>" }
      ]
    }
  ],
  "actions": [
    {
      "type": "Action.OpenUrl",
      "title": "Open in SharePoint",
      "url": "<absolute item URL>"
    }
  ]
}
```

- `buildCard(items: IShareItem[], message: string, senderName: string): IAdaptiveCardAttachment`
- Supports **single or multiple items**. For multi-select, the card body renders each item as a row in the FactSet or as separate TextBlock entries, with individual "Open in SharePoint" action buttons.
- Returns the full `chatMessage` attachment structure ready for Graph API consumption.

---

### 5. Custom Hooks

| Hook | Purpose | Key Details |
|---|---|---|
| `useTeamsDestinations` | Fetches both `/me/chats` (with timeout) and `/me/joinedTeams` in parallel | Returns `{ destinations, loading, error, chatsError }`. Fetches parallelly and supports incremental loading. If `/me/chats` times out (8s), returns empty chats, sets `chatsError: true` to display a validation warning, but allows immediate selection of Teams/Channels. |
| `useSendMessage` | Wraps the `POST` call for sending the Adaptive Card | Returns `{ send(destination, items, message), status, error }`. Manages `idle → sending → success/error` transitions. |
| `useDebounce` | Debounces a value by N ms | Generic `useDebounce<T>(value: T, delay: number): T`. |

---

### 6. Models & Interfaces

**File:** `models/interfaces.ts`

```typescript
// ─── Selected SharePoint Item ────────────────────────────
export interface IShareItem {
  id: number;
  title: string;
  fileRef: string;           // ServerRelativeUrl
  fileLeafRef: string;       // File name (for doc libs)
  listTitle: string;
  absoluteUrl: string;       // Full URL to DispForm or file
  modified: string;          // ISO date
}

// ─── Destination (unified) ───────────────────────────────
export type DestinationType = 'chat' | 'channel';

export interface IDestination {
  id: string;
  displayName: string;
  type: DestinationType;
  teamId?: string;           // Only for channels
  channelId?: string;        // Only for channels
  chatId?: string;           // Only for chats
  memberNames?: string[];    // Chat participant names
}

// ─── Teams API Models ────────────────────────────────────
export interface IChatResponse {
  id: string;
  topic: string | null;
  chatType: 'oneOnOne' | 'group' | 'meeting';
  members: IChatMember[];
}

export interface IChatMember {
  displayName: string;
  userId: string;
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

// ─── Adaptive Card ───────────────────────────────────────
export interface IAdaptiveCardAttachment {
  contentType: 'application/vnd.microsoft.card.adaptive';
  contentUrl: null;
  content: string;            // JSON-stringified card
}

// ─── Consolidated Component State ────────────────────────
export type SendStatus = 'idle' | 'sending' | 'success' | 'error';

export interface ISharePanelState {
  selectedDestination: IDestination | null;
  messageText: string;
  sendStatus: SendStatus;
  errorMessage: string;
  destinations: IDestination[];
  destinationsLoading: boolean;
  searchQuery: string;
  chatsError: boolean;
}

export const initialSharePanelState: ISharePanelState = {
  selectedDestination: null,
  messageText: '',
  sendStatus: 'idle',
  errorMessage: '',
  destinations: [],
  destinationsLoading: true,
  searchQuery: '',
  chatsError: false,
};

// ─── Component Props ─────────────────────────────────────
export interface ISharePanelProps {
  items: IShareItem[];       // Multi-select: array of selected items
  onDismiss: () => void;
}

export interface ISelectedItemsListProps {
  items: IShareItem[];
}

export interface IDestinationPickerProps {
  destinations: IDestination[];
  selectedDestination: IDestination | null;
  loading: boolean;
  searchQuery: string;
  onSearchChange: (query: string) => void;
  onDestinationSelected: (dest: IDestination) => void;
  chatsError: boolean;
}

export interface IMessageComposerProps {
  value: string;
  onChange: (value: string) => void;
  maxLength: number;
}

export interface IStatusIndicatorProps {
  status: SendStatus;
  errorMessage?: string;
  onRetry?: () => void;
}

// ─── Dialog Props ────────────────────────────────────────
export interface IShareDialogProps {
  items: IShareItem[];       // Multi-select: array of selected items
  onClose: () => void;
}
```

---

### 7. Constants

**File:** `constants/constants.ts`

```typescript
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export const GRAPH_ENDPOINTS = {
  ME_CHATS: '/me/chats?$expand=members&$top=50',
  ME_JOINED_TEAMS: '/me/joinedTeams',
  TEAM_CHANNELS: (teamId: string) => `/teams/${teamId}/channels`,
  CHAT_MESSAGES: (chatId: string) => `/chats/${chatId}/messages`,
  CHANNEL_MESSAGES: (teamId: string, channelId: string) =>
    `/teams/${teamId}/channels/${channelId}/messages`,
} as const;

export const UI = {
  PANEL_SIZE: 'medium' as const,
  DEBOUNCE_MS: 300,
  MESSAGE_MAX_LENGTH: 280,
  SUCCESS_DISMISS_MS: 3000,
  COMMAND_KEY: 'SEND_TO_TEAMS',
} as const;

export const ADAPTIVE_CARD = {
  SCHEMA: 'http://adaptivecards.io/schemas/adaptive-card.json',
  VERSION: '1.4',
  CONTENT_TYPE: 'application/vnd.microsoft.card.adaptive' as const,
} as const;

export const STRINGS = {
  COMMAND_TITLE: 'Send to Teams',
  PANEL_TITLE: 'Share via Teams Chat',
  SEARCH_PLACEHOLDER: 'Search chats, teams, or channels…',
  MESSAGE_PLACEHOLDER: 'Add an optional message…',
  SEND_BUTTON: 'Send',
  CANCEL_BUTTON: 'Cancel',
  SENDING_LABEL: 'Sending to Teams…',
  SUCCESS_MESSAGE: 'Message sent successfully!',
  ERROR_DEFAULT: 'Something went wrong. Please try again.',
  NO_RESULTS: 'No matching chats or channels found.',
  ITEM_CARD_HEADING: 'You're sharing',
} as const;
```

---

### 8. Styling Strategy

All styles use **Fluent UI v9 `makeStyles`** (Griffel). No global CSS.

**Pattern per component:**

```typescript
// *.styles.ts
import { makeStyles, tokens, shorthands } from '@fluentui/react-components';

export const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM),
  },
  // ...
});
```

- Use `tokens.*` for all colours, spacing, typography, and border-radii → ensures automatic dark/light theme support.
- Zero hard-coded pixel values for spacing; use token-based spacing.
- Each component file imports its co-located `*.styles.ts`.

---

## Configuration Files

### `config/package-solution.json` (permissions block)

```json
{
  "solution": {
    "name": "react-share-via-teams-chat-extension",
    "id": "<GUID>",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "skipFeatureDeployment": false,
    "isDomainIsolated": false,
    "developer": {
      "name": "",
      "websiteUrl": "",
      "privacyUrl": "",
      "termsOfUseUrl": "",
      "mpnId": "Undefined-1.23.0"
    },
    "webApiPermissionRequests": [
      { "resource": "Microsoft Graph", "scope": "Chat.ReadWrite" },
      { "resource": "Microsoft Graph", "scope": "Team.ReadBasic.All" },
      { "resource": "Microsoft Graph", "scope": "Channel.ReadBasic.All" },
      { "resource": "Microsoft Graph", "scope": "ChannelMessage.Send" },
      { "resource": "Microsoft Graph", "scope": "User.Read" }
    ],
    "metadata": {
      "shortDescription": { "default": "Share list items to Teams chats and channels" },
      "longDescription": { "default": "A ListView Command Set extension that lets users share selected list items or documents into a Microsoft Teams chat or channel as a formatted Adaptive Card. Supports multi-select." },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    }
  },
  "paths": {
    "zippedPackage": "solution/react-share-via-teams-chat-extension.sppkg"
  }
}
```

> [!IMPORTANT]
> `skipFeatureDeployment` is set to **`false`** — the extension must be explicitly activated on each site collection where it is needed (**site-scoped deployment**).

### `config/serve.json`

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
  "port": 4321,
  "https": true,
  "initialPage": "https://{tenantDomain}/_layouts/15/workbench.aspx"
}
```

### `elements.xml` (for debug deployment)

```xml
<?xml version="1.0" encoding="utf-8"?>
<Elements xmlns="http://schemas.microsoft.com/sharepoint/">
  <CustomAction
    Title="ShareViaTeamsChat"
    RegistrationId="100"
    RegistrationType="List"
    Location="ClientSideExtension.ListViewCommandSet.CommandBar"
    ClientSideComponentId="<component-guid>"
    ClientSideComponentProperties="{}">
  </CustomAction>
</Elements>
```

---

## Performance Considerations

| Concern | Mitigation |
|---|---|
| **Bundle size** | Import only the specific Fluent UI v9 components used (tree-shakable). Avoid importing the entire `@fluentui/react-components` barrel. |
| **Graph call volume** | Cache `chats` and `joinedTeams` results in `useRef` across re-renders within a single dialog session. No redundant fetches. |
| **Lazy channel loading** | Channels for a specific team are fetched **only when the user expands that team** in the picker — not upfront. |
| **Debounced search** | Client-side filtering uses `useDebounce(300ms)` to avoid re-renders on every keystroke. |
| **Consolidated state** | Single state object per component (spread-merge updates) reduces the number of `setState` calls and avoids intermediate renders from multiple individual state updates. |
| **Dialog teardown** | `ReactDOM.unmountComponentAtNode()` in `onAfterClose()` ensures clean disposal of the React tree and all event listeners. |
| **Minimal re-renders** | Child components are pure presentational — they receive props from `SharePanel` and never own their own state. `React.memo` applied where applicable. |
| **Strict TypeScript** | `strict: true` in `tsconfig.json` to catch null/undefined at compile time. |

---

## Execution Order

| Phase | Tasks |
|---|---|
| **1 — Scaffold** | Generate the SPFx 1.23 project (ListView Command Set template). Install `@fluentui/react-components` and `@fluentui/react-icons`. |
| **2 — Foundation** | Create `constants/`, `models/`, `loc/` folders. Define all interfaces, types, constants, and localisation strings. |
| **3 — Services** | Implement `GraphService` (singleton), `TeamsService`, and `AdaptiveCardBuilder`. |
| **4 — Hooks** | Implement `useTeamsChats`, `useJoinedTeams`, `useSendMessage`, `useDebounce`. |
| **5 — Components** | Build leaf components first (`ItemPreviewCard`, `StatusIndicator`, `MessageComposer`, `DestinationPicker`), then compose into `SharePanel`. |
| **6 — Dialog** | Create `ShareDialog.ts` (`BaseDialog`) and wire it to the React tree. |
| **7 — Command Set** | Wire `onExecute()` in `ShareViaTeamsChatCommandSet.ts` to open the dialog. |
| **8 — Config** | Update `package-solution.json` with Graph permissions. Update `elements.xml` for debug deployment. |
| **9 — Test & Polish** | Deploy to a dev tenant, approve API permissions, test end-to-end. Check light/dark theme. Verify Adaptive Card rendering in Teams. |

---

## User Experience Flow

```
┌─────────────────────────────────────────────────┐
│  SharePoint List View                           │
│  ┌───────────────────────────────────────────┐  │
│  │ ☐  Title           Modified     Author    │  │
│  │ ☑  Contoso RFP     Jun 18       J. Smith  │  │  ← User selects rows
│  │ ☑  Budget Q3       Jun 15       A. Kumar  │  │  ← Multi-select
│  └───────────────────────────────────────────┘  │
│                                                 │
│  [ Send to Teams ▸ ]   ← Button appears        │
│                                                 │
└─────────────────────────────────────────────────┘
              │
              ▼
┌─────────────────────────────────────────────────┐
│  Modal Dialog (centred overlay)                 │
│  ┌───────────────────────────────────────────┐  │
│  │  Share via Teams Chat              [ ✕ ]  │  │
│  ├───────────────────────────────────────────┤  │
│  │                                           │  │
│  │  📋 2 items selected                      │  │
│  │  ┌─────────────────────────────────────┐  │  │
│  │  │  📄 Contoso RFP · Jun 18           │  │  │
│  │  │  📄 Budget Q3   · Jun 15           │  │  │
│  │  └─────────────────────────────────────┘  │  │
│  │                                           │  │
│  │  Send to:                                 │  │
│  │  ┌─────────────────────────────────────┐  │  │
│  │  │  🔍 Search chats, teams...         │  │  │
│  │  │  ─────────────────────────────────  │  │  │
│  │  │  Recent Chats                      │  │  │
│  │  │    👤 John Smith                   │  │  │
│  │  │    👥 Project Alpha                │  │  │
│  │  │  Teams & Channels                  │  │  │
│  │  │    🏢 Engineering > General        │  │  │
│  │  │    🏢 Engineering > Design         │  │  │
│  │  └─────────────────────────────────────┘  │  │
│  │                                           │  │
│  │  Message (optional):                      │  │
│  │  ┌─────────────────────────────────────┐  │  │
│  │  │ Hey, take a look at these docs…    │  │  │
│  │  └─────────────────────────────────────┘  │  │
│  │                           236/280 chars   │  │
│  │                                           │  │
│  ├───────────────────────────────────────────┤  │
│  │          [ Cancel ]   [ Send ▸ ]          │  │
│  └───────────────────────────────────────────┘  │
└─────────────────────────────────────────────────┘
```

---

## Verification Plan

### Automated Checks

```bash
# TypeScript compilation (zero errors)
npx heft build --clean

# Linting
npx eslint src/ --ext .ts,.tsx

# Package the solution
npx heft build --clean && npx gulp bundle --ship && npx gulp package-solution --ship
```

### Manual Verification

| # | Scenario | Expected |
|---|---|---|
| 1 | Select 0 items in the list | "Send to Teams" button is **hidden** |
| 2 | Select 1 item | Button is **visible** |
| 3 | Select 3 items | Button is **visible** (multi-select supported) |
| 4 | Click "Send to Teams" with 1 item | Modal dialog opens showing single item preview |
| 5 | Click "Send to Teams" with 3 items | Modal dialog opens showing "3 items selected" with item list |
| 6 | Search "John" in picker | Filters chats/channels containing "John" |
| 7 | Select a chat + type message + click Send | Spinner shown → success toast → modal closes |
| 8 | Open Teams, navigate to the chat | Adaptive Card is visible with all shared items, links, and message |
| 9 | Click "Open in SharePoint" on the card | Navigates back to the list item or document |
| 10 | Network error during send | Error MessageBar shown with retry button |
| 11 | Toggle SharePoint dark theme | Modal and all components respect dark mode via FluentProvider |
| 12 | Activate extension on Site A, verify not on Site B | Extension only visible on sites where the feature is activated (site-scoped) |

---

## Decisions Locked In

| Decision | Resolution |
|---|---|
| **React version** | React **17.0.1** (SPFx built-in). No React 18 APIs. |
| **State management** | Single consolidated state object per component, updated via spread-merge. No individual `useState` per field. |
| **Multi-select** | ✅ Supported. `IShareItem[]` passed throughout. Adaptive Card renders all items. |
| **Deployment scope** | **Site-scoped** (`skipFeatureDeployment: false`). Must be activated per site collection. |
| **People Picker** | Not needed in core solution. User will provide a custom People Picker component if required before development starts. |
| **Destination selection** | Fluent UI `Combobox` for Teams channels, chats, and groups. |
| **UI pattern** | Centred modal `Dialog` (not side drawer). |
