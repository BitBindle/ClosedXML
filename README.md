![ClosedXML](resources/logo/readme.png)

[![CI](https://github.com/BitBindle/ClosedXML/actions/workflows/ci.yml/badge.svg?branch=develop)](https://github.com/BitBindle/ClosedXML/actions/workflows/ci.yml) [![Release](https://img.shields.io/github/v/release/BitBindle/ClosedXML?include_prereleases)](https://github.com/BitBindle/ClosedXML/releases) [![.NET](https://img.shields.io/badge/.NET-10.0-512BD4.svg)](#install)

ClosedXML is a .NET library that reads, changes and writes Excel 2007+ files (`.xlsx`, `.xlsm`). It gives an easy interface to the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). ClosedXML does not need the Excel application. For example, a web server can use it to create Excel reports.

This repository is a fork of [ClosedXML](https://github.com/ClosedXML/ClosedXML). The fork has its own releases, issues and changes. The license is MIT (see [LICENSE](LICENSE)).

## Example

```c#
using var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("Sample Sheet");
worksheet.Cell("A1").Value = "Hello World!";
worksheet.Cell("A2").FormulaA1 = "MID(A1, 7, 5)";
workbook.SaveAs("HelloWorld.xlsx");
```

## Install

The release workflow publishes the `ClosedXML` and `ClosedXML.IO` packages to the GitHub Packages feed of BitBindle. The packages target `net10.0`. ClosedXML also needs `ClosedXML.Parser` 3.0.0 or later from the same feed.

GitHub Packages needs authentication, also for a public package. Use a token with the `read:packages` scope.

nuget.org has packages with the same IDs from the original project. Use package source mapping, so that NuGet gets these packages only from the GitHub Packages feed of BitBindle. In this example `nuget.config`, NuGet reads the user name and the token from the `GITHUB_PACKAGES_USER` and `GITHUB_PACKAGES_TOKEN` environment variables:

```xml
<configuration>
  <packageSources>
    <clear />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
    <add key="github-bitbindle" value="https://nuget.pkg.github.com/BitBindle/index.json" />
  </packageSources>
  <packageSourceCredentials>
    <github-bitbindle>
      <add key="Username" value="%GITHUB_PACKAGES_USER%" />
      <add key="ClearTextPassword" value="%GITHUB_PACKAGES_TOKEN%" />
    </github-bitbindle>
  </packageSourceCredentials>
  <packageSourceMapping>
    <packageSource key="github-bitbindle">
      <package pattern="ClosedXML" />
      <package pattern="ClosedXML.IO" />
      <package pattern="ClosedXML.Parser" />
    </packageSource>
    <packageSource key="nuget.org">
      <package pattern="*" />
    </packageSource>
  </packageSourceMapping>
</configuration>
```

Then add the package to your project. The current versions are previews, so use the `--prerelease` option:

```
dotnet add package ClosedXML --prerelease
```

## Documentation

The public API is not stable. Read the release notes and the migration guide before you update.

* [Releases and release notes](https://github.com/BitBindle/ClosedXML/releases)
* [Migration guide for 0.107](docs/migrations/migrate-to-0.107.rst)
* [Documentation source](docs/index.rst)

## Frequent answers

* Column widths from `AdjustToContents` can be wrong when the computer does not have the fonts of the workbook. See [missing fonts](docs/tips/missing-font.rst).
* ClosedXML is not thread-safe. Parallel operations on one workbook can fail. The Open XML SDK is also not thread-safe.

## Contribute

Report a bug or request a feature in the [issues](https://github.com/BitBindle/ClosedXML/issues). Read the [developer guidelines](CONTRIBUTING.md) before you open a pull request.

## Credits

These people made the original ClosedXML project:

* Manuel de Leon created the project.
* Jan Havlíček, Francois Botha and Aleksei Pankratev maintained and developed it.
* Tobaloidee designed the logo.
* Many other people contributed to it.
