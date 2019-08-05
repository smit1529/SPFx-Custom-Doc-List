declare interface ICustomDocListWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TitleFieldLabel: string;
  SelectListFieldLabel: string;
  GroupByFieldLabel: string;
  SortByFieldLabel: string;
  ShoeColumnsLabel: string;
}

declare module 'CustomDocListWebPartStrings' {
  const strings: ICustomDocListWebPartStrings;
  export = strings;
}
