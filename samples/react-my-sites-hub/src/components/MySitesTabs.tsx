import {
  Tab,
  TabList,
  Overflow,
  OverflowItem,
  useOverflowMenu,
  useIsOverflowItemVisible,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Button,
  makeStyles,
  tokens,
  SelectTabData,
  SelectTabEvent,
} from "@fluentui/react-components"
import { MoreHorizontalRegular } from "@fluentui/react-icons"
import { Fade } from "@fluentui/react-motion-components-preview"
import * as React from "react"
import { Tabs } from "../utils/Strings"
import { FrequentSites } from "./FrequentSites"
import { FollowedSites } from "./FollowedSites"
import { RecentDocuments } from "./RecentDocuments"
import { FavoriteLists } from "./FavoriteLists"
// import { CustomSites } from "./CustomSites"

const TAB_COMPONENTS: Record<string, React.ReactNode> = {
  followed: <FollowedSites />,
  frequent: <FrequentSites />,
  recentDocuments: <RecentDocuments />,
  favoriteLists: <FavoriteLists />,
  // custom: <CustomSites />,
}

type TabConfig = (typeof Tabs)[number]

const useStyles = makeStyles({
  menu: {
    backgroundColor: tokens.colorNeutralBackground1,
  },
  menuButton: {
    alignSelf: "center",
  },
})

const OverflowMenuItem: React.FC<{
  tab: TabConfig
  selected: boolean
  onClick: () => void
}> = ({ tab, selected, onClick }) => {
  const isVisible = useIsOverflowItemVisible(tab.key)

  if (isVisible) {
    return null
  }

  return (
    <MenuItem key={tab.key} icon={<tab.icon filled={selected} />} onClick={onClick}>
      {tab.label}
    </MenuItem>
  )
}

const OverflowMenu: React.FC<{
  selectedTab: string
  onTabSelect: (tabId: string) => void
}> = ({ selectedTab, onTabSelect }) => {
  const { ref, isOverflowing, overflowCount } =
    useOverflowMenu<HTMLButtonElement>()
  const styles = useStyles()

  if (!isOverflowing) {
    return null
  }

  return (
    <Menu hasIcons>
      <MenuTrigger disableButtonEnhancement>
        <Button
          appearance='transparent'
          className={styles.menuButton}
          ref={ref}
          icon={<MoreHorizontalRegular />}
          aria-label={`${overflowCount} more tabs`}
          role='tab'
        />
      </MenuTrigger>
      <MenuPopover>
        <MenuList className={styles.menu}>
          {Tabs.map((tab) => (
            <OverflowMenuItem
              key={tab.key}
              tab={tab}
              selected={tab.key === selectedTab}
              onClick={() => onTabSelect(tab.key)}
            />
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  )
}

export const MySitesTabs = () => {
  const [selectedTab, setSelectedTab] = React.useState<string>("followed")

  const onTabSelect = (_: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value as string)
  }

  return (
    <>
      <Overflow minimumVisible={2}>
        <TabList
          appearance='filled-circular'
          aria-label='My Sites Tabs'
          selectedValue={selectedTab}
          onTabSelect={onTabSelect}
        >
          {Tabs.map((tab) => (
            <OverflowItem
              key={tab.key}
              id={tab.key}
              priority={tab.key === selectedTab ? 2 : 1}
            >
              <Tab value={tab.key} icon={<tab.icon filled={tab.key === selectedTab} />}>
                {tab.label}
              </Tab>
            </OverflowItem>
          ))}
          <OverflowMenu selectedTab={selectedTab} onTabSelect={setSelectedTab} />
        </TabList>
      </Overflow>
      <div
        style={{
          marginTop: "1rem",
          overflowY: "scroll",
          height: "calc(100vh - 150px)",
        }}
      >
        {/* Keyed on the active tab so switching tabs remounts and replays the
            subtle fade-in for the incoming panel. */}
        <Fade key={selectedTab} appear visible duration={200}>
          <div>{TAB_COMPONENTS[selectedTab]}</div>
        </Fade>
      </div>
    </>
  )
}
