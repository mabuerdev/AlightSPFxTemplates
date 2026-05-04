declare interface I<%= componentName.pascal %>WebPartStrings {
   // Credentials
    AlightCredentials: string;
    Author: string;
    VersionText: string;
    PropertyPaneDescription: string;
}

declare module '<%= componentName.pascal %>WebPartStrings' {
  const strings: I<%= componentName.pascal %>WebPartStrings;
  export = strings;
}
