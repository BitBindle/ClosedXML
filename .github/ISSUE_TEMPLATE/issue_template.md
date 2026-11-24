---
name: Issue
about: Issue template

---

## Read and complete the full issue template

Do not randomly delete sections. They are here for a reason.

**Do you want to request a *feature* or report a *bug*?**
- [x] Bug
- [ ] Feature
- [ ] Question

**Did you test against the latest release?**
- [ ] Yes
- [ ] No

If you answered `No`, please test with [the latest release](https://github.com/BitBindle/ClosedXML/releases) first.

**Version of ClosedXML**

e.g. 0.107.3

**What is the current behavior?**

Complete this.

**What is the expected behavior or new feature?**

Complete this.

**Is this a regression from the previous version?**

Regressions get higher priority. Test against the latest release of the previous minor version. Give the two versions that you tested. For 0.107, the previous version is 0.106 of the original project on nuget.org.

## Reproducibility
**This is an important section. Read it carefully.**

If you do not give a code sample, the developers possibly cannot reproduce the problem. Give a [minimal, complete and verifiable](https://stackoverflow.com/help/mcve) code sample. If applicable, attach a sample spreadsheet. Remove sensitive information.

**Code to reproduce problem:**
```c#
public void Main()
{
    // Code standards:
    // - Fully runnable. I should be able to copy and paste this code into a 
    //   console application and run it without having to edit it much.
    // - Declare all your variables (this follows from the previous point)
    // - The code should be a minimal code sample to illustrate issue. Don't
    //   post your full application.
}
```

- [ ] I attached a sample spreadsheet.  (You can drag files on to this issue)
