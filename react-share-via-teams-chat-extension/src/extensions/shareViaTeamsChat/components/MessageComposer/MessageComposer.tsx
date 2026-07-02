import * as React from 'react';
import { Field, Textarea } from '@fluentui/react-components';
import { STRINGS } from '../../constants/constants';
import type { IMessageComposerProps } from '../../models/interfaces';
import { useStyles } from './MessageComposer.styles';

export const MessageComposer: React.FC<IMessageComposerProps> = React.memo(
  ({ value, onChange, maxLength }) => {
    const styles = useStyles();
    const over = value.length > maxLength;

    return (
      <Field label="Message (optional)" className={styles.root}>
        <Textarea
          value={value}
          placeholder={STRINGS.MESSAGE_PLACEHOLDER}
          onChange={(_event, data) => onChange(data.value)}
          textarea={{ className: styles.textarea, maxLength: maxLength + 50 }}
          resize="vertical"
        />
        <span className={`${styles.counter} ${over ? styles.counterOver : ''}`}>
          {STRINGS.CHAR_COUNTER(value.length, maxLength)}
        </span>
      </Field>
    );
  }
);

MessageComposer.displayName = 'MessageComposer';
