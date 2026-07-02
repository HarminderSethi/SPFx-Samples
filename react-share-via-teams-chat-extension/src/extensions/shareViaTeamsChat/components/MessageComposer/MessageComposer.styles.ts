import { makeStyles, shorthands, tokens } from '@fluentui/react-components';

export const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXS)
  },
  textarea: {
    minHeight: '88px'
  },
  counter: {
    alignSelf: 'flex-end',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3
  },
  counterOver: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold
  }
});
