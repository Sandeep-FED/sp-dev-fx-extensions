import * as React from "react"
import { Slide } from "@fluentui/react-motion-components-preview"

// Subtle staggered entrance for list rows: a short fade paired with a gentle
// upward rise, cascaded by a small per-item delay so rows settle in one after
// another. The delay is capped so long lists don't wait noticeably before the
// last rows appear.
//
// The Slide presence component clones its single child and drives the animation
// through a ref (no wrapper DOM), so wrapping a <ListItem> keeps the <li> valid
// inside a Fluent <List>.
const STAGGER_STEP_MS = 40
const MAX_STAGGERED_ITEMS = 12

export interface IStaggerItemProps {
  index: number
  children: React.ReactElement
}

export const StaggerItem: React.FC<IStaggerItemProps> = ({
  index,
  children,
}) => (
  <Slide
    appear
    visible
    fromY='8px'
    duration={250}
    delay={Math.min(index, MAX_STAGGERED_ITEMS) * STAGGER_STEP_MS}
  >
    {children}
  </Slide>
)
