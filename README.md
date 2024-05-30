# Configuring App Environment for on-prem 2019

## Manage App Catalog
[Manage Catalog](https://learn.microsoft.com/en-us/sharepoint/administration/manage-the-app-catalog)

# Generating SharePoint Solutions with Yo
Choose SPO, 

 @microsoft dependencies should be ~1.4.1 in package.json file

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

# Create App Service Management proxy for APPs
# deploy assets to https://:domain:/_layouts/15
# add site collection admin
# Use Copy-Item to deploy src files to sharepoint virtual directory
