# Configuring App Environment

## Manage App Catalog
[Manage Catalog](https://learn.microsoft.com/en-us/sharepoint/administration/manage-the-app-catalog)

# Generating SharePoint Solutions with Yo
Choose 2016 and up, upgrade @microsoft dependencies to ~1.4.1 in package.json file

 "dependencies": {
    "@microsoft/sp-core-library": "~1.4.1",
    "@microsoft/sp-webpart-base": "~1.4.1",
    "@microsoft/sp-lodash-subset": "~1.4.1",
    "@microsoft/sp-office-ui-fabric-core": "~1.4.1",
    "@types/webpack-env": ">=1.12.1 <1.14.0"
  },
  "devDependencies": {
    "@microsoft/sp-build-web": "~1.4.1",
    "@microsoft/sp-module-interfaces": "~1.4.1",
    "@microsoft/sp-webpart-workbench": "~1.4.1"
  }

