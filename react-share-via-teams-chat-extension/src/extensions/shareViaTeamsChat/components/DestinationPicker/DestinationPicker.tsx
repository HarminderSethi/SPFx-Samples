import * as React from 'react';
import {
  Avatar,
  Button,
  Field,
  Spinner,
  Tag,
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  TagPickerOptionGroup,
  Text
} from '@fluentui/react-components';
import type { TagPickerInputProps, TagPickerProps } from '@fluentui/react-components';
import { People24Regular } from '@fluentui/react-icons';
import { STRINGS, UI } from '../../constants/constants';
import { useDebounce } from '../../hooks/useDebounce';
import type {
  IDestination,
  IDestinationPickerProps
} from '../../models/interfaces';
import { useStyles } from './DestinationPicker.styles';

const filterDestinations = (list: IDestination[], query: string): IDestination[] => {
  const q = query.trim().toLowerCase();
  if (!q) return list;
  return list.filter((d) =>
    [d.displayName, ...(d.memberEmails ?? [])].some((value) =>
      value.toLowerCase().includes(q)
    )
  );
};

const renderOption = (
  styles: ReturnType<typeof useStyles>,
  destination: IDestination
): JSX.Element => {
  const secondaryText = destination.type === 'chat'
    ? (destination.memberEmails?.filter(Boolean).join(', ') || destination.memberNames?.filter(Boolean).join(', ') || undefined)
    : undefined;

  return (
    <TagPickerOption
      key={destination.id}
      value={destination.id}
      text={destination.displayName}
      media={destination.type === 'chat' ? <Avatar size={24} name={destination.displayName} /> : <People24Regular />}
      secondaryContent={secondaryText}
    >
      <div className={styles.optionContent}>
        <div className={styles.optionText}>
          <Text className={styles.optionTitle}>{destination.displayName}</Text>
        </div>
      </div>
    </TagPickerOption>
  );
};

export const DestinationPicker: React.FC<IDestinationPickerProps> = React.memo((props) => {
  const {
    destinations,
    selectedDestinations,
    loading,
    searchQuery,
    onSearchChange,
    onDestinationsSelected,
    chatsError
  } = props;
  const styles = useStyles();
  const debouncedQuery = useDebounce(searchQuery, UI.DEBOUNCE_MS);
  const selectedOptionIds = React.useMemo(
    () => selectedDestinations.map((destination) => destination.id),
    [selectedDestinations]
  );

  const { chats, channels } = React.useMemo(() => {
    const filtered = filterDestinations(destinations, debouncedQuery).filter(
      (destination) => !selectedOptionIds.includes(destination.id)
    );
    return {
      chats: filtered.filter((d) => d.type === 'chat'),
      channels: filtered.filter((d) => d.type === 'channel')
    };
  }, [destinations, debouncedQuery, selectedOptionIds]);

  const handleInput: TagPickerInputProps['onInput'] = (event) => {
    const target = event.target as HTMLInputElement;
    onSearchChange(target.value);
  };

  const handleOptionSelect: TagPickerProps['onOptionSelect'] = (_event, data) => {
    const nextDestinations = data.selectedOptions
      .map((id) => destinations.find((destination) => destination.id === id))
      .filter((destination): destination is IDestination => Boolean(destination));
    onDestinationsSelected(nextDestinations);
    onSearchChange('');
  };

  const handleClear = (): void => {
    onDestinationsSelected([]);
    onSearchChange('');
  };

  const hasResults = chats.length + channels.length > 0;

  return (
    <Field
      label="Send to"
      className={styles.root}
      validationState={chatsError ? 'warning' : undefined}
      validationMessage={chatsError ? 'Recent chats could not be loaded due to a timeout. Teams and channels are still available.' : undefined}
    >
      <TagPicker
        selectedOptions={selectedOptionIds}
        onOptionSelect={handleOptionSelect}
        appearance="outline"
        size="large"
      >
        <TagPickerControl
          className={styles.pickerControl}
          expandIcon={loading ? <Spinner size="tiny" /> : undefined}
          secondaryAction={selectedDestinations.length > 0 ? (
            <Button
              appearance="transparent"
              className={styles.clearButton}
              size="small"
              onClick={handleClear}
            >
              {STRINGS.CLEAR_DESTINATIONS}
            </Button>
          ) : undefined}
        >
          <TagPickerGroup>
            {selectedDestinations.map((destination) => (
              <Tag
                key={destination.id}
                shape="rounded"
                media={destination.type === 'chat' ? <Avatar name={destination.displayName} /> : <People24Regular />}
                value={destination.id}
                size="medium"
              >
                {destination.displayName}
              </Tag>
            ))}
          </TagPickerGroup>
          <TagPickerInput
            className={styles.picker}
            clearable
            placeholder={STRINGS.SEARCH_PLACEHOLDER}
            value={searchQuery}
            onInput={handleInput}
          />
        </TagPickerControl>
        <TagPickerList>
          {!hasResults && !loading && (
            <div className={styles.emptyState}>
              <div>{STRINGS.NO_RESULTS}</div>
              {destinations.length === 0 && (
                <div className={styles.hintText}>{STRINGS.CHAT_PERMISSION_HINT}</div>
              )}
            </div>
          )}
          {chats.length > 0 && (
            <TagPickerOptionGroup label={STRINGS.RECENT_CHATS_GROUP}>
              {chats.map((d) => renderOption(styles, d))}
            </TagPickerOptionGroup>
          )}
          {channels.length > 0 && (
            <TagPickerOptionGroup label={STRINGS.TEAMS_CHANNELS_GROUP}>
              {channels.map((d) => renderOption(styles, d))}
            </TagPickerOptionGroup>
          )}
        </TagPickerList>
      </TagPicker>
    </Field>
  );
});

DestinationPicker.displayName = 'DestinationPicker';
