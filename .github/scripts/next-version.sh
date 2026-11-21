#!/usr/bin/env bash
# Prints the next version to release.
#
# Major and minor version are defined in Directory.Build.props (VersionMajorMinor), the version suffix
# (e.g. "preview") in VersionSuffix. The patch version is the highest patch of an existing tag
# <major>.<minor>.<patch>[-suffix] plus one, or 0 when there is no such tag.
#
# Usage: next-version.sh            reads the tags of the current repository
#        next-version.sh <tags...>  uses the passed tags (for testing)
set -euo pipefail

project="ClosedXML/ClosedXML.csproj"
major_minor="$(dotnet msbuild "$project" -getProperty:VersionMajorMinor | tr -d '[:space:]')"
suffix="$(dotnet msbuild "$project" -getProperty:VersionSuffix | tr -d '[:space:]')"

if [[ -z "$major_minor" ]]; then
  echo "VersionMajorMinor is not defined in Directory.Build.props" >&2
  exit 1
fi

if [[ $# -gt 0 ]]; then tags=("$@"); else mapfile -t tags < <(git tag --list "${major_minor}.*"); fi

pattern="^${major_minor//./\\.}\\.([0-9]+)(-.+)?$"
patch=-1
for tag in "${tags[@]}"; do
  if [[ "$tag" =~ $pattern ]] && (( BASH_REMATCH[1] > patch )); then
    patch="${BASH_REMATCH[1]}"
  fi
done

version="${major_minor}.$((patch + 1))"
[[ -n "$suffix" ]] && version="${version}-${suffix}"
echo "$version"
