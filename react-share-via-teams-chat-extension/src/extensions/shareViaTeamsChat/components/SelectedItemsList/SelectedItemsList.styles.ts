import { makeStyles, shorthands, tokens } from '@fluentui/react-components';

export const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalS),
    maxHeight: '260px',
    overflowY: 'auto',
    ...shorthands.padding(tokens.spacingVerticalXS)
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS),
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300
  },
  countBadge: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
    ...shorthands.borderRadius(tokens.borderRadiusCircular),
    ...shorthands.padding('2px', tokens.spacingHorizontalS),
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold
  },
  list: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalXS)
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap(tokens.spacingHorizontalS),
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalS),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius(tokens.borderRadiusMedium)
  },
  rowText: {
    display: 'flex',
    flexDirection: 'column',
    minWidth: 0,
    flexGrow: 1
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  meta: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  icon: {
    flexShrink: 0,
    color: tokens.colorBrandForeground1
  },
  removeButton: {
    minWidth: '28px',
    width: '28px',
    height: '28px',
    color: tokens.colorNeutralForeground2,
    fontWeight: tokens.fontWeightSemibold,
    flexShrink: 0,
    ':hover': {
      color: tokens.colorPaletteRedForeground1,
      backgroundColor: tokens.colorPaletteRedBackground1
    }
  },
  emptyState: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalS),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius(tokens.borderRadiusMedium)
  }
});
