import { Text, List, ListItem, Divider, Link } from "@fluentui/react-components"
import * as React from "react"
import { getFavoriteLists } from "../services/SiteService"
import { AppContext } from "./MySitesRoot"
import { BaseComponentContext } from "@microsoft/sp-component-base"
import { IRecentList } from "../models/IRecentList"
import { useEffect, useState } from "react"
import { ListBarColor } from "@fluentui/react-icons"
import { SiteListShimmer } from "./Shimmer"
import { StaggerItem } from "./StaggerItem"
import { useListItemStyles } from "./useListItemStyles"

export const FavoriteLists = () => {
  const context = React.useContext(AppContext) as BaseComponentContext
  const styles = useListItemStyles()

  const [favoriteLists, setFavoriteLists] = useState<IRecentList[]>([])
  const [loading, setLoading] = useState<boolean>(true)

  useEffect(() => {
    favoriteListsFetch()
  }, [context])

  const favoriteListsFetch = async () => {
    try {
      setLoading(true)
      const response = await getFavoriteLists(context)
      setFavoriteLists(response)
    } catch (error) {
      console.error("Error fetching favorite lists:", error)
    } finally {
      setLoading(false)
    }
  }

  const buildListUrl = (list: IRecentList): string => {
    if (/^https?:\/\//i.test(list.Url)) {
      return list.Url
    }
    return `${context.pageContext.web.absoluteUrl.replace(
      context.pageContext.web.serverRelativeUrl,
      "",
    )}${list.Url}`
  }

  if (loading) {
    return <SiteListShimmer count={5} />
  }

  return (
    <List navigationMode='items'>
      {favoriteLists.map((list: IRecentList, index) => (
        <React.Fragment key={list.Id}>
          <StaggerItem index={index}>
            <ListItem
              value={list.Title}
              data-value={list.Title}
              aria-label={list.Title}
              className={styles.row}
              style={{
                paddingBottom: "1rem",
                display: "flex",
                gap: "10px",
                alignItems: "center",
              }}
            >
              <ListBarColor fontSize={34} />
              <Link
                href={buildListUrl(list)}
                target='_blank'
                style={{ textDecoration: "none", color: "inherit" }}
              >
                <div style={{ display: "flex", flexDirection: "column" }}>
                  <Text size={400}>{list.Title}</Text>
                  {(list.SiteTitle || list.WebTitle) && (
                    <Text size={200} style={{ opacity: 0.7 }}>
                      {list.SiteTitle || list.WebTitle}
                    </Text>
                  )}
                </div>
              </Link>
            </ListItem>
          </StaggerItem>
          {index < favoriteLists.length - 1 && (
            <Divider style={{ paddingBottom: "1rem" }} />
          )}
        </React.Fragment>
      ))}
    </List>
  )
}
