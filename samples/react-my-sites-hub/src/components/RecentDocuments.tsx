import { Text, List, ListItem, Divider, Link } from "@fluentui/react-components"
import * as React from "react"
import { getRecentDocuments } from "../services/SiteService"
import { AppContext } from "./MySitesRoot"
import { BaseComponentContext } from "@microsoft/sp-component-base"
import { IRecentDocument } from "../models/IRecentDocument"
import { useEffect, useState } from "react"
import { DocumentColor } from "@fluentui/react-icons"
import { SiteListShimmer } from "./Shimmer"
import { StaggerItem } from "./StaggerItem"
import { useListItemStyles } from "./useListItemStyles"

export const RecentDocuments = () => {
  const context = React.useContext(AppContext) as BaseComponentContext
  const styles = useListItemStyles()

  const [recentDocuments, setRecentDocuments] = useState<IRecentDocument[]>([])
  const [loading, setLoading] = useState<boolean>(true)

  useEffect(() => {
    recentDocumentsFetch()
  }, [context])

  const recentDocumentsFetch = async () => {
    try {
      setLoading(true)
      const response = await getRecentDocuments(context)
      setRecentDocuments(response)
    } catch (error) {
      console.error("Error fetching recent documents:", error)
    } finally {
      setLoading(false)
    }
  }

  if (loading) {
    return <SiteListShimmer count={5} />
  }

  return (
    <List navigationMode='items'>
      {recentDocuments.map((doc: IRecentDocument, index) => (
        <React.Fragment key={doc.Id}>
          <StaggerItem index={index}>
            <ListItem
              value={doc.Title}
              data-value={doc.Title}
              aria-label={doc.Title}
              className={styles.row}
              style={{
                paddingBottom: "1rem",
                display: "flex",
                gap: "10px",
                alignItems: "center",
              }}
            >
              <DocumentColor fontSize={34} />
              <Link
                href={doc.Url}
                target='_blank'
                style={{ textDecoration: "none", color: "inherit" }}
              >
                <div style={{ display: "flex", flexDirection: "column" }}>
                  <Text size={400}>{doc.Title}</Text>
                  {doc.ContainerName && (
                    <Text size={200} style={{ opacity: 0.7 }}>
                      {doc.ContainerName}
                    </Text>
                  )}
                </div>
              </Link>
            </ListItem>
          </StaggerItem>
          {index < recentDocuments.length - 1 && (
            <Divider style={{ paddingBottom: "1rem" }} />
          )}
        </React.Fragment>
      ))}
    </List>
  )
}
