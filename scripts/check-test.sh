#!/usr/bin/env bash
# Tests scripts/check.sh. Each fixture line has a known result.
# The test runs check.sh in a temporary git repository, so the fixture is not in this repository.
# Usage: scripts/check-test.sh
set -euo pipefail

script="$(cd "$(dirname "$0")" && pwd)/check.sh"
work="$(mktemp -d)"
trap 'rm -rf "$work"' EXIT

cd "$work"
git init --quiet
git -c user.name=test -c user.email=test@example.com commit --quiet --allow-empty -m base
mkdir Tests

# Lines that check.sh must report. CRLF line endings are the same as in the C# files.
printf '%s\r\n' \
    '// #1742 skip the locale prefix' \
    '// Fixes #1234.' \
    '// See #12, and more' \
    '/* Fix for issue 955 */' \
    '/// <a href="https://github.com/o/r/issues/7">#7</a>' \
    '// Codeplex #6932' \
    'public void Issue_1916_CanMeasure() {}' \
    'public void Issue1917NotContainsFilter() {}' \
    'public async Task CanSave_ForGH123() {}' \
    '[TestCase(1, TestName = "Fix #12")]' \
    '[TestCase(2, TestName = "Fix #13")] // The value is negative.' \
    '// See GH-1234 for the cause.' \
    > Tests/Report.cs

# Lines that check.sh must not report.
printf '%s\r\n' \
    '// See https://example.com/a//b for the format.' \
    '// The color is #00ffcc.' \
    '/// <c>&#931;Values</c>' \
    '// ECMA-376 Part 1, 18.3.1.73 requires this.' \
    '//        3 '"'"'#,##0'"'"'' \
    '//       48 '"'"'##0.0E+0'"'"' and #0.00' \
    '// Wait until task 2 completes.' \
    'public void DebugMode2() {}' \
    'public void CanReadTable() {}' \
    > Tests/Pass.cs

failures=0
output="$("$script" HEAD)" && status=0 || status=$?

if [ "$status" -ne 1 ]; then
    echo "FAIL: exit code is $status, expected 1."
    failures=$((failures + 1))
fi

for line in $(seq 1 12); do
    if ! grep -q "^Tests/Report.cs:$line: " <<< "$output"; then
        echo "FAIL: check.sh does not report Tests/Report.cs:$line."
        failures=$((failures + 1))
    fi
done

if grep "^Tests/Pass.cs:" <<< "$output"; then
    echo "FAIL: check.sh reports the lines above, but they are correct."
    failures=$((failures + 1))
fi

rm Tests/Report.cs Tests/Pass.cs
if ! "$script" HEAD > /dev/null; then
    echo "FAIL: check.sh fails when the change has no lines."
    failures=$((failures + 1))
fi

if "$script" no-such-base 2> /dev/null; then status=0; else status=$?; fi
if [ "$status" -ne 2 ]; then
    echo "FAIL: exit code for a base that does not exist is $status, expected 2."
    failures=$((failures + 1))
fi

if [ "$failures" -gt 0 ]; then
    echo "check-test.sh: $failures failure(s)."
    exit 1
fi
echo "check-test.sh: all tests pass."
