# Site Content Ribbon Extension

A lightweight SharePoint Framework application customizer that adds a ribbon-style “Site content” button to the SharePoint header and opens a modern, searchable experience for browsing site content directly from the page.

This extension is designed to help users quickly discover and navigate lists, document libraries, and other site assets without leaving the current page context.

## Overview

The extension injects a new button into the SharePoint header area and opens a Fluent UI drawer panel that presents site content in a tabular experience. It uses SharePoint app tile data to surface content such as lists, libraries, and apps, and provides a more discoverable entry point for site exploration.

## Features

- Adds a header-mounted “Site content” button for quick access
- Opens a right-side drawer panel with a polished, modern UI
- Displays site content items with useful metadata such as type, item count, modified date, and description
- Supports live search filtering to quickly find the required content
- Supports sorting by column to improve content discovery
- Includes context actions for content items
- Uses Microsoft Graph to retrieve Drive details for document libraries when needed

## Screenshots

![Ribbon button](sharepoint/assets/site%20content%20ribbon%20button.png)

![Site content panel](sharepoint/assets/site%20content%20panel.png)

![Search and filtering](sharepoint/assets/site%20content%20filter.png)

![Item menu actions](sharepoint/assets/site%20content%20item%20menu%20options.png)

## Technology Stack

- SharePoint Framework (SPFx)
- React 17
- Fluent UI React components
- SharePoint REST API via SPHttpClient
- Microsoft Graph API via MSGraphClientV3

## Prerequisites

Before using this solution, make sure you have:

- A SharePoint Online tenant with permissions to deploy SPFx solutions
- Node.js 22.x (the project is configured for Node 22.14.x)
- A development environment capable of building SPFx solutions
- Appropriate permissions to access site content and Microsoft Graph resources used by the extension

## Getting Started

### Install dependencies

```bash
npm install
```

### Run locally in debug mode

```bash
npm run start
```

### Build for packaging

```bash
npm run build
```

## How it works

The customizer waits for the SharePoint header region to become available and mounts the ribbon button into the header. When the user clicks the button, the extension queries SharePoint app tiles and renders them in a drawer-based UI for fast browsing and search.

## Project Structure

- src/extensions/siteContentRibbon/components – UI components for the ribbon button and drawer experience
- src/extensions/siteContentRibbon/services – SharePoint and Microsoft Graph integration
- sharepoint/assets – screenshots and deployment assets

## License

This project is provided as-is for learning and customization purposes. Please review and adjust for your own environment and governance requirements.

## Disclaimer

THIS CODE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.