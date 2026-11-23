using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Comments;

[TestFixture]
internal class XLCommentTests
{
    [Test]
    [Issue("123")]
    public void New_line_of_comment_is_a_line_feed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var comment = ws.Cell("A1").CreateComment();

        comment.AddText("a");
        comment.AddNewLine();
        comment.AddText("b");

        Assert.AreEqual("a\nb", comment.Text);
    }

    [Test]
    public void Copy_of_cell_copies_text_and_style_of_comment()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var source = ws.Cell("B2");
        var comment = source.CreateComment();
        comment.AddText("Bold").SetBold();
        comment.AddText(" text");
        comment.Style.ColorsAndLines.SetFillColor(XLColor.Red);

        var target = ws.Cell("D5").CopyFrom(source);

        Assert.True(target.HasComment);
        var copy = target.GetComment();
        Assert.AreNotSame(comment, copy);
        Assert.AreEqual("Bold text", copy.Text);
        Assert.True(copy.First().Bold);
        Assert.False(copy.Last().Bold);
        Assert.AreEqual(XLColor.Red, copy.Style.ColorsAndLines.FillColor);
        Assert.AreNotEqual(comment.ShapeId, copy.ShapeId);
        Assert.False(copy.HasPhonetics);
    }

    [Test]
    public void Copy_of_cell_copies_phonetics_of_comment()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var source = ws.Cell("B2");
        var comment = source.CreateComment();
        comment.AddText("東京");
        comment.Phonetics
            .SetFontSize(7)
            .SetType(XLPhoneticType.Hiragana)
            .SetAlignment(XLPhoneticAlignment.Center)
            .Add("とうきょう", 0, 2);

        var target = ws.Cell("D5").CopyFrom(source);

        var copy = target.GetComment();
        Assert.True(copy.HasPhonetics);
        Assert.AreEqual(1, copy.Phonetics.Count);
        var phonetic = copy.Phonetics.Single();
        Assert.AreEqual("とうきょう", phonetic.Text);
        Assert.AreEqual(0, phonetic.Start);
        Assert.AreEqual(2, phonetic.End);
        Assert.AreEqual(7, copy.Phonetics.FontSize);
        Assert.AreEqual(XLPhoneticType.Hiragana, copy.Phonetics.Type);
        Assert.AreEqual(XLPhoneticAlignment.Center, copy.Phonetics.Alignment);
    }

    [Test]
    public void Delete_removes_comment_from_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("B2");
        cell.CreateComment().AddText("Note");

        cell.GetComment().Delete();

        Assert.False(cell.HasComment);
    }
}
