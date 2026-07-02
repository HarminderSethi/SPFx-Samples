import { ADAPTIVE_CARD } from '../constants/constants';
import type {
  IAdaptiveCardAttachment,
  IChatMessageBody,
  IShareItem
} from '../models/interfaces';

interface IAdaptiveCardElement {
  type: string;
  [key: string]: unknown;
}

interface IAdaptiveCardAction {
  type: string;
  title: string;
  url?: string;
}

interface IAdaptiveCard {
  type: 'AdaptiveCard';
  $schema: string;
  version: string;
  body: IAdaptiveCardElement[];
  actions?: IAdaptiveCardAction[];
}

const formatDate = (iso: string): string => {
  if (!iso) return '';
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return iso;
  return date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};


const buildHeader = (): IAdaptiveCardElement => ({
  type: 'ColumnSet',
  columns: [
    {
      type: 'Column',
      width: 'stretch',
      items: [
        {
          type: 'TextBlock',
          text: 'Shared from SharePoint',
          weight: 'Bolder',
          size: 'Medium',
          wrap: true
        }
      ]
    }
  ]
});

const buildItemBlock = (item: IShareItem, index: number, total: number): IAdaptiveCardElement[] => {
  const blocks: IAdaptiveCardElement[] = [];
  if (total > 1) {
    blocks.push({
      type: 'TextBlock',
      text: `${index + 1}. ${item.title}`,
      weight: 'Bolder',
      size: 'Medium',
      wrap: true,
      spacing: index === 0 ? 'Medium' : 'Small'
    });
  } else {
    blocks.push({
      type: 'TextBlock',
      text: item.title,
      weight: 'Bolder',
      size: 'Large',
      wrap: true,
      spacing: 'Medium'
    });
  }
  const factTitle = item.isLibrary ? 'Library' : 'List';
  blocks.push({
    type: 'FactSet',
    facts: [
      { title: factTitle, value: item.listTitle },
      { title: 'Modified', value: formatDate(item.modified) }
    ]
  });
  return blocks;
};

const buildItemsSummaryTable = (items: IShareItem[]): IAdaptiveCardElement => ({
  type: 'Container',
  spacing: 'Medium',
  items: [
    {
      type: 'TextBlock',
      text: '',
      weight: 'Bolder',
      wrap: true,
      spacing: 'None'
    },
    {
      type: 'Table',
      firstRowAsHeader: true,
      showGridLines: true,
      columns: [
        { width: 3.1 },
        { width: 1.6 },
        { width: 1.6 },
        { width: 0.9 }
      ],
      rows: [
        {
          type: 'TableRow',
          cells: [
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: 'Item',
                  weight: 'Bolder',
                  horizontalAlignment: 'Center',
                  wrap: true
                }
              ]
            },
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: 'Type',
                  weight: 'Bolder',
                  horizontalAlignment: 'Center',
                  wrap: true
                }
              ]
            },
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: 'Modified',
                  weight: 'Bolder',
                  horizontalAlignment: 'Center',
                  wrap: true
                }
              ]
            },
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: 'Link',
                  weight: 'Bolder',
                  horizontalAlignment: 'Center',
                  wrap: true
                }
              ]
            }
          ]
        },
        ...items.map((item) => ({
          type: 'TableRow',
          cells: [
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: item.title,
                  wrap: true,
                  horizontalAlignment: 'Left',
                  isSubtle: true
                }
              ]
            },
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: item.isLibrary ? 'Document' : 'List item',
                  wrap: true,
                  horizontalAlignment: 'Left'
                }
              ]
            },
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: formatDate(item.modified),
                  wrap: true,
                  horizontalAlignment: 'Left'
                }
              ]
            },
            {
              type: 'TableCell',
              items: [
                {
                  type: 'TextBlock',
                  text: `[Open](${item.absoluteUrl})`,
                  wrap: true,
                  color: 'Accent',
                  horizontalAlignment: 'Left'
                }
              ]
            }
          ]
        }))
      ]
    }
  ]
});

export class AdaptiveCardBuilder {
  public static buildCard(
    items: IShareItem[],
    message: string
  ): IAdaptiveCardAttachment {
    if (!items.length) {
      throw new Error('AdaptiveCardBuilder.buildCard requires at least one item.');
    }

    const body: IAdaptiveCardElement[] = [buildHeader()];

    if (message.trim().length > 0) {
      body.push({
        type: 'TextBlock',
        text: message,
        wrap: true,
        spacing: 'Medium'
      });
    }

    if (items.length > 1) {
      body.push(buildItemsSummaryTable(items));
    } else {
      buildItemBlock(items[0], 0, 1).forEach((block) => body.push(block));
    }

    const actions: IAdaptiveCardAction[] = items.length === 1
      ? [{ type: 'Action.OpenUrl', title: 'Open in SharePoint', url: items[0].absoluteUrl }]
      : [];

    const card: IAdaptiveCard = {
      type: 'AdaptiveCard',
      $schema: ADAPTIVE_CARD.SCHEMA,
      version: ADAPTIVE_CARD.VERSION,
      body,
      actions
    };

    return {
      id: '1',
      contentType: ADAPTIVE_CARD.CONTENT_TYPE,
      contentUrl: undefined,
      content: JSON.stringify(card)
    };
  }

  public static buildMessageBody(
    items: IShareItem[],
    message: string
  ): IChatMessageBody {
    const attachment = AdaptiveCardBuilder.buildCard(items, message);
    return {
      body: {
        contentType: 'html',
        content: `<attachment id="${attachment.id}"></attachment>`
      },
      attachments: [attachment]
    };
  }
}
