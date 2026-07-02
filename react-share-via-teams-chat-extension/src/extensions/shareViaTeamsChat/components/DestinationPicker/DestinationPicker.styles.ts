import { makeStyles, shorthands, tokens } from '@fluentui/react-components';

export const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXS)
  },
  label: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  picker: {
    width: '100%'
  },
  pickerControl: {
    width: '100%'
  },
  clearButton: {
    minWidth: 'auto'
  },
  optionContent: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS)
  },
  optionText: {
    display: 'flex',
    flexDirection: 'column',
    minWidth: 0
  },
  optionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  optionMeta: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3
  },
  emptyState: {
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalS),
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXS)
  },
  hintText: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200
  }
});
