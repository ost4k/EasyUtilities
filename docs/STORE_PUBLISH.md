# EasyUtilities - Microsoft Store publish

## 1) Store description dependency disclosure (required)
In Partner Center, add the dependency notice within the first two lines of the Store description.

Use this exact opening:

```text
Requires Microsoft .NET 8.0 Desktop Runtime (x64).
If it is not installed, install it first, then run EasyUtilities.
```

Then continue with the rest of your product description text.

## 2) Update Store identity
Before final Store upload, replace identity placeholders in `EasyUtilities/Package.appxmanifest`:

- `Identity Name`
- `Identity Publisher`
- `Properties/PublisherDisplayName`

Use values from your Partner Center app reservation.

## 3) Build Store upload package (MSIX upload)
From repo root:

```powershell
dotnet publish "EasyUtilities\\EasyUtilities.csproj" `
  -c Release `
  -p:UapAppxPackageBuildMode=StoreUpload `
  -p:AppxBundle=Always `
  -p:AppxBundlePlatforms=x64 `
  -p:AppxPackageSigningEnabled=false `
  -p:GenerateAppxPackageOnBuild=true
```

## 4) Output location
Package artifacts are written to:

`EasyUtilities\\AppPackages\\`

Upload the generated `.msixupload` (or Store-upload bundle artifact) to Partner Center.

## 5) Notes
- Local sideload install requires signing; Store upload does not.
- Keep app version in manifest in sync with release.
