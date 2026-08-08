// Constants.ts
import {
  bundleIcon,
  GlobeStarFilled,
  GlobeStarRegular,
  ClockFilled,
  ClockRegular,
  DocumentFilled,
  DocumentRegular,
  ListBarFilled,
  ListBarRegular,
} from "@fluentui/react-icons"

// Bundled icons render the filled glyph when their `filled` prop is set (used
// for the active tab) and the regular glyph otherwise.
const FollowedIcon = bundleIcon(GlobeStarFilled, GlobeStarRegular)
const FrequentIcon = bundleIcon(ClockFilled, ClockRegular)
const RecentDocumentsIcon = bundleIcon(DocumentFilled, DocumentRegular)
const FavoriteListsIcon = bundleIcon(
  ListBarFilled,
  ListBarRegular
)

export const DrawerTitles = {
  mySitesPanel: "My Sites Hub",
  aiInsightsPanel: "AI Insights",
}

export const Tooltips = {
  followedTab: "Sites you are following",
  frequentTab: "Sites you visit frequently",
  recommendedTab: "Recommended sites for you",
  followButton: "Follow this site",
  unfollowButton: "Unfollow this site",
  pinButton: "Pin panel",
  closeButton: "Close panel",
}

export const Tabs = [
  { key: "followed", label: "Followed", icon: FollowedIcon },
  { key: "frequent", label: "Frequent", icon: FrequentIcon },
  {
    key: "recentDocuments",
    label: "Recent Documents",
    icon: RecentDocumentsIcon,
  },
  {
    key: "favoriteLists",
    label: "Favorite Lists",
    icon: FavoriteListsIcon,
  },
]
