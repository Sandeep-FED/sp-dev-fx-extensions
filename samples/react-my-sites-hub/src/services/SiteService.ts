import { BaseComponentContext } from "@microsoft/sp-component-base"
import {
  SPHttpClient,
  SPHttpClientResponse,
  MSGraphClientV3,
} from "@microsoft/sp-http"
import { IFollowedSites } from "../models/IFollowedSites";
import { IRecentList } from "../models/IRecentList";
import { IRecentDocument } from "../models/IRecentDocument";

export interface IFollowedSitesResponse {
  value: IFollowedSites[]
}

export const getFollowedSites = async (
  context: BaseComponentContext
): Promise<IFollowedSites[]> => {
  try {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/v2.1/favorites/followedSites?$top=30`

    const response: SPHttpClientResponse = await context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json',
          'Caller': 'WebAppBar',
          'Prefer': 'fillSiteCount=30'
        }
      }
    )

    if (!response.ok) {
      throw new Error(`Error fetching followed sites: ${response.statusText}`)
    }

    const data: IFollowedSitesResponse = await response.json()
    return data.value
  } catch (error) {
    console.error("Error in getFollowedSites:", error)
    throw error
  }
}

// GetGroupImage only exists for Microsoft 365 group-connected sites.
// Communication sites and non-group team sites have no backing group, so the
// call fails (or returns 200 with an empty/non-image body). Treat any of those
// as "no image" and let the caller fall back to the site acronym rather than
// surfacing an error or rendering a broken image.
export const getGroupImageUrl = async (
  context: BaseComponentContext,
  siteUrl: string
): Promise<string> => {
  try {
    const endpoint = `${siteUrl}/_api/GroupService/GetGroupImage`

    const response: SPHttpClientResponse = await context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )

    if (!response.ok) {
      return ""
    }

    const blob = await response.blob()
    // A site with no group image can still respond 200 with an empty or
    // non-image payload; guard against building a broken blob URL from it.
    if (blob.size === 0 || !blob.type.startsWith("image/")) {
      return ""
    }

    return URL.createObjectURL(blob)
  } catch (error) {
    console.error("Error in getGroupImageUrl:", error)
    return ""
  }
}

// get frequent sites
export const getFrequentSites = async (
  context: BaseComponentContext
): Promise<IFollowedSites[]> => {
  try {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/v2.1/insights/frequentSites?$top=30`

    const response: SPHttpClientResponse = await context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    )

    if (!response.ok) {
      throw new Error(`Error fetching frequent sites: ${response.statusText}`)
    }

    const data: IFollowedSitesResponse = await response.json()
    return data.value
  } catch (error) {
    console.error("Error in getFrequentSites:", error)
    throw error
  }
}

// HTTP/2 responses often have an empty statusText, so build a useful error
// message from the status code plus the response body.
const readResponseError = async (
  response: SPHttpClientResponse
): Promise<string> => {
  let body = ""
  try {
    body = await response.text()
  } catch {
    // ignore body read failures
  }
  return `${response.status} ${response.statusText} ${body}`.trim()
}

// get the current user's recently used documents via Microsoft Graph.
// SharePoint's GetRecentLists only runs on the personal (-my) site, which an
// app customizer on a team site can't reach, so this tab uses /me/insights/used
// instead. It returns recent documents across M365 with rich metadata (title,
// file type, and the site/container that holds them). Requires the delegated
// Graph permission Sites.Read.All (see config/package-solution.json).
export const getRecentDocuments = async (
  context: BaseComponentContext
): Promise<IRecentDocument[]> => {
  try {
    const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient(
      "3"
    )

    const response = await client.api("/me/insights/used").top(30).get()

    const items: any[] = response?.value ?? []

    return items
      .filter((item) => item?.resourceReference?.webUrl)
      .map((item) => {
        const visualization = item.resourceVisualization ?? {}
        const reference = item.resourceReference ?? {}
        const lastUsed = item.lastUsed ?? {}
        return {
          Id: item.id ?? reference.id ?? "",
          Title: visualization.title ?? "Untitled",
          Url: reference.webUrl,
          Type: visualization.type,
          ContainerName: visualization.containerDisplayName,
          PreviewImageUrl: visualization.previewImageUrl,
          LastUsed:
            lastUsed.lastAccessedDateTime ?? lastUsed.lastModifiedDateTime,
        } as IRecentDocument
      })
  } catch (error) {
    console.error("Error in getRecentDocuments:", error)
    throw error
  }
}

// The list endpoints can return the collection under different envelopes
// (value / d.results / d.<Method>.results) and with varying property casing
// depending on the OData metadata level. Normalize both here.
const normalizeRecentLists = (raw: unknown): IRecentList[] => {
  const container = raw as Record<string, any>
  const d = container?.d
  // In verbose mode the array is nested under d.<MethodName>.results, whose key
  // varies by endpoint (GetRecentLists, GetListsBy, ...), so find it generically.
  const nestedResults =
    d && typeof d === "object"
      ? (Object.values(d).find(
          (v: any) => v && Array.isArray(v.results)
        ) as any)?.results
      : undefined

  const items: unknown[] =
    (Array.isArray(raw) && (raw as unknown[])) ||
    container?.value ||
    d?.results ||
    nestedResults ||
    []

  const pick = (item: Record<string, any>, keys: string[]): any => {
    for (const key of keys) {
      const match = Object.keys(item).find(
        (k) => k.toLowerCase() === key.toLowerCase()
      )
      if (match !== undefined && item[match] !== undefined) {
        return item[match]
      }
    }
    return undefined
  }

  return items.map((entry) => {
    const item = entry as Record<string, any>
    return {
      Id: pick(item, ["Id", "ListId", "id"]) ?? "",
      Title: pick(item, ["Title", "Name", "DisplayName"]) ?? "",
      Url:
        pick(item, ["listUrl", "Url", "ServerRelativeUrl", "WebUrl", "Path"]) ??
        "",
      BaseTemplate: pick(item, [
        "BaseTemplate",
        "TemplateType",
        "ListTemplateType",
      ]),
      IsDocLib: pick(item, ["isDocLib", "IsDocLib"]),
      ListId: pick(item, ["ListId"]),
      WebId: pick(item, ["WebId"]),
      SiteId: pick(item, ["SiteId"]),
      SiteTitle: pick(item, ["SiteTitle"]),
      WebTitle: pick(item, ["WebTitle"]),
      IconUrl: pick(item, ["IconUrl"]),
      LastActivityOn: pick(item, ["LastActivityOn", "LastActivityTime"]),
    } as IRecentList
  })
}

// get favorite (saved) lists for the current user
export const getFavoriteLists = async (
  context: BaseComponentContext
): Promise<IRecentList[]> => {
  try {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/SP.FavoriteLists.GetListsBy`

    const response: SPHttpClientResponse = await context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
        },
      }
    )

    if (!response.ok) {
      throw new Error(
        `Error fetching favorite lists: ${await readResponseError(response)}`
      )
    }

    const raw = await response.json()
    // Log the raw payload once so the response shape can be confirmed per tenant.
    console.log("getFavoriteLists raw response", raw)

    return normalizeRecentLists(raw)
  } catch (error) {
    console.error("Error in getFavoriteLists:", error)
    throw error
  }
}

export const addFollowedSite = async (
  context: BaseComponentContext,
  siteId: string,
  siteUrl: string
): Promise<void> => {
  try {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/v2.1/favorites/followedSites/oneDrive.add`

     const requestBody = {
      value: [
        {
          webUrl: siteUrl,
          id: siteId,
        },
      ],
    }

    const response: SPHttpClientResponse = await context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify(requestBody),
        headers: {
          "Content-Type": "application/json",
        },
      }
    )

    if (!response.ok) {
      throw new Error(`Error adding followed site: ${response.statusText}`)
    }
  } catch (error) {
    console.error("Error in addFollowedSite:", error)
    throw error
  }
}

// remove followed site _api/v2.1/favorites/followedSites/oneDrive.remove
export const removeFollowedSite = async (
  context: BaseComponentContext,
  webUrl: string,
  siteId: string
): Promise<void> => {
  try {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/v2.1/favorites/followedSites/oneDrive.remove`

    const requestBody = {
      value: [
        {
          webUrl: webUrl,
          id: siteId,
        },
      ],
    }

    const response: SPHttpClientResponse = await context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify(requestBody),
        headers: {
          "Content-Type": "application/json",
        },
      }
    )

    if (!response.ok) {
      throw new Error(`Error removing followed site: ${response.statusText}`)
    }
  } catch (error) {
    console.error("Error in removeFollowedSite:", error)
    throw error
  }
}