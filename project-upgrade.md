# Upgrade project react-user-custom-action-editor to v1.14.0

Date: 3/21/2022

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.14.0. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i -SE @microsoft/sp-core-library@1.14.0
```

File: [./package.json:13:9](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i -SE @microsoft/sp-lodash-subset@1.14.0
```

File: [./package.json:14:9](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i -SE @microsoft/sp-office-ui-fabric-core@1.14.0
```

File: [./package.json:15:9](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-webpart-base@1.14.0
```

File: [./package.json:17:9](./package.json)

### FN001021 @microsoft/sp-property-pane | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-property-pane

Execute the following command:

```sh
npm i -SE @microsoft/sp-property-pane@1.14.0
```

File: [./package.json:16:9](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i -DE @microsoft/sp-build-web@1.14.0
```

File: [./package.json:31:9](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i -DE @microsoft/sp-module-interfaces@1.14.0
```

File: [./package.json:32:9](./package.json)

### FN002009 @microsoft/sp-tslint-rules | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-tslint-rules

Execute the following command:

```sh
npm i -DE @microsoft/sp-tslint-rules@1.14.0
```

File: [./package.json:33:9](./package.json)

### FN006005 package-solution.json metadata | Required

In package-solution.json add metadata section

```json
{
  "solution": {
    "metadata": {
      "shortDescription": {
        "default": "react-user-custom-action-editor description"
      },
      "longDescription": {
        "default": "react-user-custom-action-editor description"
      },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    }
  }
}
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN006006 package-solution.json features | Required

In package-solution.json add features section

```json
{
  "solution": {
    "features": [
      {
        "title": "react-user-custom-action-editor Feature",
        "description": "The feature that activates elements of the react-user-custom-action-editor solution.",
        "id": "467c6ba8-05b2-4793-b961-8a42908b44ff",
        "version": "1.0.0.0"
      }
    ]
  }
}
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.14.0"
  }
}
```

File: [./.yo-rc.json:5:9](./.yo-rc.json)

### FN014008 Hosted workbench type in .vscode/launch.json | Recommended

In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration

```json
{
  "configurations": [
    {
      "type": "pwa-chrome"
    }
  ]
}
```

File: [.vscode\launch.json:12:7](.vscode\launch.json)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm i -SE @microsoft/sp-core-library@1.14.0 @microsoft/sp-lodash-subset@1.14.0 @microsoft/sp-office-ui-fabric-core@1.14.0 @microsoft/sp-webpart-base@1.14.0 @microsoft/sp-property-pane@1.14.0
npm i -DE @microsoft/sp-build-web@1.14.0 @microsoft/sp-module-interfaces@1.14.0 @microsoft/sp-tslint-rules@1.14.0
npm dedupe
```

### Modify files

#### [./config/package-solution.json](./config/package-solution.json)

In package-solution.json add metadata section:

```json
{
  "solution": {
    "metadata": {
      "shortDescription": {
        "default": "react-user-custom-action-editor description"
      },
      "longDescription": {
        "default": "react-user-custom-action-editor description"
      },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    }
  }
}
```

In package-solution.json add features section:

```json
{
  "solution": {
    "features": [
      {
        "title": "react-user-custom-action-editor Feature",
        "description": "The feature that activates elements of the react-user-custom-action-editor solution.",
        "id": "467c6ba8-05b2-4793-b961-8a42908b44ff",
        "version": "1.0.0.0"
      }
    ]
  }
}
```

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.14.0"
  }
}
```

#### [.vscode\launch.json](.vscode\launch.json)

In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration:

```json
{
  "configurations": [
    {
      "type": "pwa-chrome"
    }
  ]
}
```
