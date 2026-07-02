import { makeStyles, shorthands, tokens } from '@fluentui/react-components';

export const useStyles = makeStyles({
  surface: {
    minWidth: 'min(560px, 90vw)',
    maxWidth: '720px'
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM),
    ...shorthands.padding(tokens.spacingVerticalS, '0')
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXS)
  },
  sendingState: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS),
    color: tokens.colorNeutralForeground2
  }
});
