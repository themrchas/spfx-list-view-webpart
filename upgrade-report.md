# Upgrade project Branding Item View to v1.4.0

Date: 11/25/2024

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.4.0. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i -SE @microsoft/sp-core-library@1.4.0
```

File: [./package.json:17:5](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i -SE @microsoft/sp-lodash-subset@1.4.0
```

File: [./package.json:18:5](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i -SE @microsoft/sp-office-ui-fabric-core@1.4.0
```

File: [./package.json:19:5](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-webpart-base@1.4.0
```

File: [./package.json:22:5](./package.json)

### FN001008 react | Required

Upgrade SharePoint Framework dependency package react

Execute the following command:

```sh
npm i -SE react@15.6.2
```

File: [./package.json:32:5](./package.json)

### FN001009 react-dom | Required

Upgrade SharePoint Framework dependency package react-dom

Execute the following command:

```sh
npm i -SE react-dom@15.6.2
```

File: [./package.json:33:5](./package.json)

### FN001005 @types/react | Required

Upgrade SharePoint Framework dependency package @types/react

Execute the following command:

```sh
npm i -SE @types/react@15.6.6
```

File: [./package.json:30:5](./package.json)

### FN001006 @types/react-dom | Required

Upgrade SharePoint Framework dependency package @types/react-dom

Execute the following command:

```sh
npm i -SE @types/react-dom@15.5.6
```

File: [./package.json:31:5](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i -DE @microsoft/sp-build-web@1.4.0
```

File: [./package.json:37:5](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i -DE @microsoft/sp-module-interfaces@1.4.0
```

File: [./package.json:38:5](./package.json)

### FN002003 @microsoft/sp-webpart-workbench | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench

Execute the following command:

```sh
npm i -DE @microsoft/sp-webpart-workbench@1.4.0
```

File: [./package.json:39:5](./package.json)

### FN006002 package-solution.json includeClientSideAssets | Required

Update package-solution.json includeClientSideAssets

```json
{
  "solution": {
    "includeClientSideAssets": true
  }
}
```

File: [./config/package-solution.json:3:15](./config/package-solution.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.4.0"
  }
}
```

File: [./.yo-rc.json:15:5](./.yo-rc.json)

### FN012005 tsconfig.json typeRoots ./node_modules/@microsoft | Required

Add ./node_modules/@microsoft to typeRoots in tsconfig.json

```json
{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@microsoft"
    ]
  }
}
```

File: [./tsconfig.json:11:5](./tsconfig.json)

### FN002007 ajv | Required

Upgrade SharePoint Framework dev dependency package ajv

Execute the following command:

```sh
npm i -DE ajv@5.2.2
```

File: [./package.json:44:5](./package.json)

### FN003001 config.json schema | Required

Update config.json schema URL

```json   Possible issue *****
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/config.2.0.schema.json"
}
```

File: [./config/config.json:1:1](./config/config.json)

### FN003002 config.json version | Required

Update config.json version number

```json
{
  "version": "2.0"
}
```

File: [./config/config.json:1:1](./config/config.json)

### FN003003 config.json bundles | Required

In config.json add the 'bundles' property

```json
{
  "bundles": {
    "branding-item-view-web-part": {
      "components": [
        {
          "entrypoint": "./lib/webparts/brandingItemView/BrandingItemViewWebPart.js",
          "manifest": "./src/webparts/brandingItemView/BrandingItemViewWebPart.manifest.json"
        }
      ]
    }
  }
}
```

File: [./config/config.json:1:1](./config/config.json)

### FN003004 config.json entries | Required

Remove the "entries" property in ./config/config.json

```json
{
  "entries": [
    {
      "entry": "./lib/webparts/brandingItemView/BrandingItemViewWebPart.js",
      "manifest": "./src/webparts/brandingItemView/BrandingItemViewWebPart.manifest.json",
      "outputPath": "./dist/branding-item-view-web-part.js"
    }
  ]
}
```

File: [./config/config.json:2:3](./config/config.json)

### FN003005 Update path of the localized resource | Required

In the config.json file, update the path of the localized resource

```json
{
  "localizedResources": {
    "BrandingItemViewWebPartStrings": "lib/webparts/brandingItemView/loc/{locale}.js"
  }
}
```

File: [./config/config.json:11:5](./config/config.json)

### FN004001 copy-assets.json schema | Required

Update copy-assets.json schema URL

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/copy-assets.schema.json"
}
```

File: [./config/copy-assets.json:2:3](./config/copy-assets.json)

### FN005001 deploy-azure-storage.json schema | Required

Update deploy-azure-storage.json schema URL

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/deploy-azure-storage.schema.json"
}
```

File: [./config/deploy-azure-storage.json:2:3](./config/deploy-azure-storage.json)

### FN006001 package-solution.json schema | Required

Update package-solution.json schema URL

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/package-solution.schema.json"
}
```

File: [./config/package-solution.json:2:3](./config/package-solution.json)

### FN007001 serve.json schema | Required

Update serve.json schema URL

```json
{
  "$schema": "https://dev.office.com/json-schemas/core-build/serve.schema.json"
}
```

File: [./config/serve.json:2:3](./config/serve.json)

### FN009001 write-manifests.json schema | Required

Update write-manifests.json schema URL

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/write-manifests.schema.json"
}
```

File: [./config/write-manifests.json:2:3](./config/write-manifests.json)

### FN011001 Web part manifest schema | Required

Update schema in manifest

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

File: [src\webparts\brandingItemView\BrandingItemViewWebPart.manifest.json:2:3](src\webparts\brandingItemView\BrandingItemViewWebPart.manifest.json)

### FN017001 Run npm dedupe | Optional

Ignored ******

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm i -SE @microsoft/sp-core-library@1.4.0 @microsoft/sp-lodash-subset@1.4.0 @microsoft/sp-office-ui-fabric-core@1.4.0 @microsoft/sp-webpart-base@1.4.0 react@15.6.2 react-dom@15.6.2 @types/react@15.6.6 @types/react-dom@15.5.6
npm i -DE @microsoft/sp-build-web@1.4.0 @microsoft/sp-module-interfaces@1.4.0 @microsoft/sp-webpart-workbench@1.4.0 ajv@5.2.2
npm dedupe
```

### Modify files

#### [./config/package-solution.json](./config/package-solution.json)

Update package-solution.json includeClientSideAssets:

```json
{
  "solution": {
    "includeClientSideAssets": true
  }
}
```

Update package-solution.json schema URL:

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/package-solution.schema.json"
}
```

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.4.0"
  }
}
```

#### [./tsconfig.json](./tsconfig.json)

Add ./node_modules/@microsoft to typeRoots in tsconfig.json:

```json
{
  "compilerOptions": {
    "typeRoots": [
      "./node_modules/@microsoft"
    ]
  }
}
```

#### [./config/config.json](./config/config.json)

Update config.json schema URL:

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/config.2.0.schema.json"
}
```

Update config.json version number:

```json
{
  "version": "2.0"
}
```

In config.json add the 'bundles' property:

```json
{
  "bundles": {
    "branding-item-view-web-part": {
      "components": [
        {
          "entrypoint": "./lib/webparts/brandingItemView/BrandingItemViewWebPart.js",
          "manifest": "./src/webparts/brandingItemView/BrandingItemViewWebPart.manifest.json"
        }
      ]
    }
  }
}
```

Remove the "entries" property in ./config/config.json:

```json
{
  "entries": [
    {
      "entry": "./lib/webparts/brandingItemView/BrandingItemViewWebPart.js",
      "manifest": "./src/webparts/brandingItemView/BrandingItemViewWebPart.manifest.json",
      "outputPath": "./dist/branding-item-view-web-part.js"
    }
  ]
}
```

In the config.json file, update the path of the localized resource:

```json
{
  "localizedResources": {
    "BrandingItemViewWebPartStrings": "lib/webparts/brandingItemView/loc/{locale}.js"
  }
}
```

#### [./config/copy-assets.json](./config/copy-assets.json)

Update copy-assets.json schema URL:

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/copy-assets.schema.json"
}
```

#### [./config/deploy-azure-storage.json](./config/deploy-azure-storage.json)

Update deploy-azure-storage.json schema URL:

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/deploy-azure-storage.schema.json"
}
```

#### [./config/serve.json](./config/serve.json)

Update serve.json schema URL:

```json
{
  "$schema": "https://dev.office.com/json-schemas/core-build/serve.schema.json"
}
```

#### [./config/write-manifests.json](./config/write-manifests.json)

Update write-manifests.json schema URL:

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx-build/write-manifests.schema.json"
}
```

#### [src\webparts\brandingItemView\BrandingItemViewWebPart.manifest.json](src\webparts\brandingItemView\BrandingItemViewWebPart.manifest.json)

Update schema in manifest:

```json
{
  "$schema": "https://dev.office.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```
