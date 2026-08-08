import { makeStyles, tokens } from "@fluentui/react-components"

// Subtle hover / press feedback for list rows. We animate only background-color
// and box-shadow here — NOT transform — because rows are wrapped in the <Slide>
// entrance motion, whose finished animation persists with `fill: forwards` and
// would override any CSS transform on the same element. Background + elevation
// give a gentle "lift" on hover and a "pressed" look on click without fighting
// the entrance animation.
export const useListItemStyles = makeStyles({
  row: {
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    transitionProperty: "background-color, box-shadow",
    transitionDuration: tokens.durationNormal,
    transitionTimingFunction: tokens.curveEasyEase,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      boxShadow: tokens.shadow4,
    },
    ":active": {
      backgroundColor: tokens.colorNeutralBackground1Pressed,
      boxShadow: tokens.shadow2,
    },
    "@media (prefers-reduced-motion: reduce)": {
      transitionDuration: "0.01ms",
    },
  },
})
