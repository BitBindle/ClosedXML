# Instructions for agents

These rules apply to each agent that changes this repository.
`CONTRIBUTING.md` gives the general developer guidelines. If a rule here is different, this file applies.

## Repository

ClosedXML is a .NET library that reads and writes Excel files (`.xlsx`, `.xlsm`).

| Project | Content |
| --- | --- |
| `ClosedXML` | The library. Targets `net10.0`. |
| `ClosedXML.IO` | Low-level readers and writers for the package parts. Targets `net10.0`. |
| `ClosedXML.IO.CodeGen` | Code generator for `ClosedXML.IO`. |
| `ClosedXML.Tests` | NUnit tests. Targets `net10.0`. |
| `ClosedXML.Examples` | Examples. Some tests use them. |
| `ClosedXML.Sandbox` | Console application for manual experiments. |

The base branch is `develop`.

## Build, test and check

Do these steps before the review gate. The commands are based on `.github/workflows/ci.yml`.

1. Build the solution:

   ```bash
   dotnet build ClosedXML.slnx --configuration Release
   ```

2. Run the tests with coverage. On Linux, skip the `OLEDB` category, because it needs a Windows driver:

   ```bash
   dotnet test ClosedXML.Tests/ClosedXML.Tests.csproj --configuration Release --no-build --filter "TestCategory!=OLEDB" --collect "Code Coverage;Format=cobertura" --results-directory test-results
   ```

3. Measure the line coverage of the new code from the Cobertura report. For example, use `diff-cover` against `origin/develop`.
4. Run the text check:

   ```bash
   scripts/check.sh
   ```

   If you change `scripts/check.sh`, also run `scripts/check-test.sh`.

5. If a step fails, fix the cause. Do not skip, disable or ignore a test to make a step pass.

## Review gate

A read-only subagent reviews each change before the change goes back to the person you work with.

1. Complete the change and the steps in "Build, test and check".
2. Save the diff against the base branch to a file outside the repository.
3. Start a new read-only reviewer subagent. In Claude Code, use the `change-reviewer` agent in `.claude/agents/change-reviewer.md`.
4. Give the reviewer the task and the path of the diff file.
5. Give the reviewer the results of the build, the tests, the coverage and `scripts/check.sh`.
6. Do not give the reviewer your own assessment.
7. The reviewer returns one verdict: **BLOCK**, **APPROVE WITH NOTES** or **APPROVE**.
8. If the verdict is **BLOCK**, fix each objection. Then do the gate again with a new reviewer.
9. If the verdict is **APPROVE WITH NOTES**, fix the notes that you can. Tell the person about the notes that remain.
10. Give the change to the person with the last verdict.

Do not give a person a change that has a **BLOCK** verdict. Do not attach the objections as caveats instead of a fix.

The reviewer uses these verdicts:

| Verdict | Condition |
| --- | --- |
| **BLOCK** | The change has a correctness bug or a security problem. Or the build, a test or `scripts/check.sh` fails. Or new code has 80% line coverage or less. Or the change breaks a rule in this file. |
| **APPROVE WITH NOTES** | The change is correct. The reviewer has minor points that do not stop the change. |
| **APPROVE** | The reviewer has no points. |

## Problems that the change did not cause

You can find a problem that the change did not cause.

1. If the fix is small and its effect is clear, fix it in the same change. Tell the reviewer and the person about it.
2. If not, log a separate GitHub issue. Give the problem, the location and the steps to see it. Do not fix it in the change.

## Code rules

### Tests

- New code has more than 80% line coverage from tests.
- Existing code does not have a coverage requirement. If the person asks for coverage of existing code, the requirement applies to that code too.
- Link a test to a GitHub issue with the `[Issue("1234")]` attribute. Do not put the issue number in the test name.

### Memory and allocations

- Use `Span<T>`, `ReadOnlySpan<T>`, `Memory<T>`, `stackalloc` and `ArrayPool<T>` where they remove an allocation or a copy on a hot path. Examples of hot paths are parsers, converters and serializers.
- Do not use them where they only make the code harder to read.

### Async code

- Code is async from end to end.
- Do not block on a task. This prohibits `.Result`, `.Wait()` and `GetAwaiter().GetResult()`.
- Do not use synchronous I/O.
- Pass a `CancellationToken` to each async call that accepts one.
- A controller action takes a `CancellationToken` as a parameter.
- The existing synchronous public API stays as it is. Change it only if the task asks for it.

### Unnecessary work

Do no unnecessary work:

- No extra round trips.
- No extra copies.
- No reads of data that the request does not use.

## Language and documentation

Use [ASD-STE-100](https://www.asd-ste100.org/) (Simplified Technical English) everywhere. This includes replies to the person you work with, code, comments, documentation, commit messages and pull requests.

- One idea per sentence. Keep instructions to 20 words and descriptions to 25. Keep a paragraph to six sentences.
- Use active voice and the present tense. Name who or what does the action.
- One word, one meaning. Choose a term and keep it. Do not change it for style.
- Use plain, common words. Do not use metaphors, idioms, literary phrases, marketing adjectives or emoji.
- Do not use jargon that the reader did not use first. "Add a tooltip to the heat glosses" is better than "the glosses find their voice at last".
- Identifiers, test names and headings follow the same rules.
- Say a thing one time, and say it directly. Do not use openers such as "it is worth noting that".
- Do not assess your own work with words such as "comprehensive", "robust" or "seamless". Do not hedge without information.
- If you can delete a sentence and lose nothing, delete it.
- Write a procedure as numbered steps in the order of the work. Put the condition before the action: "If the build fails, read the policy output."
- Comments and test titles give the reason or the spec requirement, not a task number. `scripts/check.sh` checks this for the lines that a change adds.
