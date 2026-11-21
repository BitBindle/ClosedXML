#!/usr/bin/env bash
# Finds task references in the C# lines that the current change adds.
#
# A comment or a test title gives the reason or the spec requirement, not a task number.
# A test links to a GitHub issue with the [Issue("1234")] attribute.
# The script checks only added lines, so existing code does not fail the check.
# The change is the difference between the merge base with <base> and the working tree.
# New files that git does not track yet are part of the change.
#
# Limit: the script reads each line alone. It does not see a block comment line without a
# leading "*".
#
# Usage: scripts/check.sh [base]    The default base is origin/develop.
# The exit code is 0 if the script finds no problem, 1 if it finds a problem, 2 on a usage error.
set -euo pipefail

cd "$(git rev-parse --show-toplevel)"

base="${1:-origin/develop}"
if ! git rev-parse --verify --quiet "$base^{commit}" > /dev/null; then
    echo "check.sh: base '$base' does not exist. Fetch it or give a different base." >&2
    exit 2
fi
if ! merge_base="$(git merge-base "$base" HEAD)"; then
    echo "check.sh: '$base' and HEAD have no common commit. In a shallow clone, fetch more history." >&2
    exit 2
fi

changes() {
    git diff -U0 --no-color --no-ext-diff "$merge_base" -- '*.cs'
    git ls-files --others --exclude-standard -z -- '*.cs' |
        while IFS= read -r -d '' file; do
            # git diff --no-index returns 1 when the files are different.
            git diff -U0 --no-color --no-ext-diff --no-index /dev/null "$file" || true
        done
}

changes | awk '
# Returns the comment text of a line, or "" if the line has no comment.
function comment_of(s,    pos, i, at) {
    if (s ~ /^[ \t]*(\/\*|\*)/)
        return s
    pos = 1
    while ((i = index(substr(s, pos), "//")) > 0) {
        at = pos + i - 1
        # "://" is part of a URL, not the start of a comment.
        if (at == 1 || substr(s, at - 1, 1) != ":")
            return substr(s, at + 2)
        pos = at + 2
    }
    i = index(s, "/*")
    if (i > 0)
        return substr(s, i + 2)
    return ""
}

# Returns 1 if the text refers to a task, an issue or a pull request by number.
function has_task_ref(s,    l) {
    l = tolower(s)
    # These are not tasks: "&#931;" (XML character reference), "#00ffcc" (color),
    # "##0.0E+0" and "#0.00" (Excel number formats). "Fixes #12." is a task.
    if (l ~ /(^|[^&#0-9a-z])#[0-9]+([^0-9a-z.,%]|[.,]([^0-9]|$)|$)/)
        return 1
    if (l ~ /(^|[^a-z])(issue|bug|gh|ticket|pr|pull request|work item|codeplex)[ _:#-]*[0-9]+/)
        return 1
    return 0
}

# Returns 1 if a test method name contains a task number, for example Issue_1916_Name.
function has_task_name(name) {
    if (name ~ /^(issue|bug|gh|ticket)_?[0-9]/)
        return 1
    if (name ~ /(^|_|[a-z0-9])(Issue|ISSUE|Bug|BUG|Gh|GH|Ticket|TICKET)_?[0-9]/)
        return 1
    return 0
}

function report(message, text) {
    printf "%s:%d: %s: %s\n", file, line, message, text
    problems++
}

/^\+\+\+ / {
    file = substr($0, 7)
    next
}

/^@@ / {
    match($0, /\+[0-9]+/)
    line = substr($0, RSTART + 1, RLENGTH - 1) + 0
    next
}

/^\+/ {
    text = substr($0, 2)
    sub(/\r$/, "", text)

    comment = comment_of(text)
    # comment_of returns the end of the line, so the code is the part before it.
    code = substr(text, 1, length(text) - length(comment))
    if (comment != "" && has_task_ref(comment))
        report("comment refers to a task number, give the reason instead", text)

    if (file ~ /Tests?\//) {
        if (match(text, /(void|Task)[ \t]+[A-Za-z0-9_]+[ \t]*\(/)) {
            name = substr(text, RSTART, RLENGTH)
            sub(/^(void|Task)[ \t]+/, "", name)
            sub(/[ \t]*\($/, "", name)
            if (has_task_name(name))
                report("test name contains a task number, use [Issue(\"...\")] instead", text)
        }
        if (code ~ /(TestName|Description)[ \t]*=/ && has_task_ref(code))
            report("test title refers to a task number, give the reason instead", text)
    }

    line++
}

END {
    if (problems > 0) {
        printf "check.sh: %d problem(s).\n", problems
        exit 1
    }
}
'
