declare interface IBirthdayWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  NumberUpComingDaysLabel: string;
  BackgroundImageLabel: string;
}

declare module 'BirthdaysWebPartStrings' {
  const strings: IBirthdayWebPartStrings;
  export = strings;
}
