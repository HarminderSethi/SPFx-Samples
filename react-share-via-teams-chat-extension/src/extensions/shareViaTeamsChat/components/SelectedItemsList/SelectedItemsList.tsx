import * as React from 'react';
import { Button, Text } from '@fluentui/react-components';
import { Document24Regular } from '@fluentui/react-icons';
import { STRINGS } from '../../constants/constants';
import type { ISelectedItemsListProps, IShareItem } from '../../models/interfaces';
import { useStyles } from './SelectedItemsList.styles';

const formatDate = (iso: string): string => {
  if (!iso) return '';
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return iso;
  return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
};

const ItemRow: React.FC<{ item: IShareItem; onRemoveItem?: (itemId: number) => void }> = ({
  item,
  onRemoveItem
}) => {
  const styles = useStyles();
  return (
    <div className={styles.row}>
      <Document24Regular className={styles.icon} />
      <div className={styles.rowText}>
        <Text className={styles.title} title={item.title}>
          {item.title}
        </Text>
        <Text className={styles.meta}>
          {item.listTitle}
          {item.modified ? ` · ${formatDate(item.modified)}` : ''}
        </Text>
      </div>
      {onRemoveItem && (
        <Button
          appearance="transparent"
          size="small"
          className={styles.removeButton}
          aria-label={`Remove ${item.title}`}
          onClick={() => onRemoveItem(item.id)}
        >
          x
        </Button>
      )}
    </div>
  );
};

export const SelectedItemsList: React.FC<ISelectedItemsListProps> = React.memo(({ items, onRemoveItem }) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <span>{STRINGS.ITEM_CARD_HEADING}</span>
        <span className={styles.countBadge}>{STRINGS.ITEMS_SELECTED(items.length)}</span>
      </div>
      <div className={styles.list}>
        {items.length === 0 ? (
          <Text className={styles.emptyState}>{STRINGS.NO_ITEMS_SELECTED}</Text>
        ) : (
          items.map((item) => (
            <ItemRow key={`${item.listTitle}-${item.id}`} item={item} onRemoveItem={onRemoveItem} />
          ))
        )}
      </div>
    </div>
  );
});

SelectedItemsList.displayName = 'SelectedItemsList';
