export interface IRecentList {
  Id: string
  Title: string
  Url: string
  BaseTemplate?: number
  IsDocLib?: boolean
  ListId?: string
  WebId?: string
  SiteId?: string
  SiteTitle?: string
  WebTitle?: string
  IconUrl?: string
  LastActivityOn?: string
}

export interface IRecentListsResponse {
  value: IRecentList[]
}
