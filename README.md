# SP_Dev — Coalition Product Catalog for SharePoint 2019

A suite of SharePoint Framework (SPFx) web parts that turn a flat SharePoint document
library into a **metadata-driven product catalog**. Built for **SharePoint 2019
on-prem** and deployed across coalition enclaves so US and allied users can find
products fast, by the metadata that matters to them, instead of scrolling folders.

Every product carries a small, consistent taxonomy (Intel Category, Involved Nations,
publish date). Once the metadata is consistent, findability is just a query, and the
web parts here are the queries: upload-and-tag, category cards, search and filter,
recent products, and an inline document viewer.

## How it works

```
Products (flat dump)  ──tag on upload──▶  Product Library (.stp)  ──REST──▶  SPFx web parts (React)  ──▶  US + Allied users
                                          • Intel Categories                  • Upload + tag
                                          • Involved Nations                  • Category cards
                                          • Publish Date                      • Search + filter
                                                                              • Recent + inline view
```

The data layer is a single SharePoint list/library defined by `ProductLibrary.stp`.
Each item exposes the fields the web parts read (internal names shown):

| Field | Internal name | Purpose |
| --- | --- | --- |
| Title | `Title` | Product name |
| Intel Categories | `Intel_x0020_Categories` | Primary taxonomy used for cards + filtering |
| Involved Nations | `Involved_x0020_Nations` | Coalition scoping / filtering |
| Publish Date | `PublishDate` | Sort + "recent products" |
| File name | `FileLeafRef` | Underlying document |
| Embed URL | `ServerRedirectedEmbedUrl` | Inline (in-browser) viewing |

## Web parts

Each folder is a self-contained SPFx solution (React + TypeScript + Office UI Fabric).

| Folder | Web part | What it does |
| --- | --- | --- |
| `UploadProducts` | Upload Products | Add a product and tag it with the required metadata as it is uploaded. |
| `CategorizeProducts` | Categorize Products | Organize and group products by Intel Category. |
| `ProductSearch` | Product Search | Filter the library by category and involved nation. |
| `ProductSearchCards` | Product Search Cards | Category cards plus a results list with an inline document viewer panel. |
| `RecentProducts` | Recent Products | Surface the newest products by publish date. |
| `DirectoryListing` | Directory Listing | Site / library navigation. |
| `HelpfulLinks` | Helpful Links | Quick links across the site. |
| `AdvancedProductSearch` | Advanced Product Search | Scaffold for an expanded search experience (starter, not feature-complete). |

`ProductLibrary.stp` — the SharePoint list template that defines the catalog schema.
Provision it in any enclave to get the identical structure the web parts expect.

## Toolchain (pinned for SharePoint 2019)

SharePoint 2019 on-prem supports the SPFx 1.4.x generation. Keep the toolchain on
those versions or the bundle will not load on-prem:

- SPFx **1.4.1**, React **15.6.x**, Office UI Fabric React **5.x**
- Node.js LTS compatible with that SPFx generation, Gulp CLI

When scaffolding or updating a web part with the Yeoman generator, confirm the
`@microsoft/sp-*` dependencies are `~1.4.1` in `package.json`:

```json
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
```

## Build a web part

From any web-part folder:

```bash
npm install
gulp serve                       # local workbench during development
gulp bundle --ship               # production bundle
gulp package-solution --ship     # produces the .sppkg in ./sharepoint/solution
```

## Deploy to SharePoint 2019 on-prem

1. **App catalog** — ensure a site collection app catalog exists.
   See [Manage the app catalog](https://learn.microsoft.com/en-us/sharepoint/administration/manage-the-app-catalog).
2. **Provision the library** — create the catalog library from `ProductLibrary.stp`
   so the metadata columns (Intel Categories, Involved Nations, Publish Date) exist.
3. **Create an App Service Management proxy** for the apps.
4. **Deploy the assets** to the SharePoint layouts directory
   (`https://<domain>/_layouts/15`). Use `Copy-Item` to push the bundled `dist`
   assets into the SharePoint virtual directory.
5. **Add the deploying account as a site collection admin.**
6. **Upload the `.sppkg`** to the app catalog, then add the app to the target site.
7. **Allow inline PDF viewing** so products open in the browser (this is not a STIG
   finding):

   ```powershell
   $webApp = Get-SPWebApplication("https://site")
   $webApp.AllowedInlineDownloadedMimeTypes.Add("application/pdf")
   $webApp.Update()
   ```

   Or use the [Central Admin panel](https://www.c-sharpcorner.com/article/enable-pdf-files-in-sharepoint-to-open-up-in-the-browser/).

## Context

Built at GDIT (USAFRICOM J26/9) as part of modernizing SharePoint across all
enclaves: a better user experience that organized products by relevant metadata so
US and allied nations could find what they needed quickly. Deployed in production
on SharePoint 2019 on-prem.
