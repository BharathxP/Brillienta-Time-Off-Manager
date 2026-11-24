declare interface ITimeOffManagerWebPartStrings {
  PropertyPaneDescription: string;
  ViewGroupName: string;
  ViewLabel: string;
  ViewApp: string;
  ViewCalendar: string;
  ViewMy: string;
  SiteGroupName: string;
  SiteUrlLabel: string;
  SiteUrlDescription: string;
  SiteUrlInvalid: string;
}

declare module 'TimeOffManagerWebPartStrings' {
  const strings: ITimeOffManagerWebPartStrings;
  export = strings;
}
