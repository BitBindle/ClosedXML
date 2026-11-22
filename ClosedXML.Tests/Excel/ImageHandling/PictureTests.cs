using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Point = System.Drawing.Point;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class PictureTests
    {
        [Test]
        public void CanAddPictureFromStream()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
                {
                    var picture = ws.AddPicture(resourceStream, "MyPicture")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(50, 50)
                        .WithSize(200, 200);

                    Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                    Assert.AreEqual(200, picture.Width);
                    Assert.AreEqual(200, picture.Height);
                }
            }
        }

        [Test]
        public void CanAddPictureFromFile()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

                try
                {
                    using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
                    using (var fileStream = File.Create(path))
                    {
                        resourceStream.Seek(0, SeekOrigin.Begin);
                        resourceStream.CopyTo(fileStream);
                        fileStream.Close();
                    }

                    var picture = ws.AddPicture(path)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(50, 50);

                    Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                    Assert.AreEqual(400, picture.Width);
                    Assert.AreEqual(400, picture.Height);
                }
                finally
                {
                    if (File.Exists(path))
                        File.Delete(path);
                }
            }
        }

        [Test]
        public void CanAddPictureConcurrentlyFromFile()
        {
            var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

            try
            {
                using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
                using (var fileStream = File.Create(path))
                {
                    resourceStream.Seek(0, SeekOrigin.Begin);
                    resourceStream.CopyTo(fileStream);
                    fileStream.Close();
                }

                Parallel.Invoke(() => verifyAddImageFromFile(path), () => verifyAddImageFromFile(path));
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        private void verifyAddImageFromFile(string filePath)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var picture = ws.AddPicture(filePath)
                           .WithPlacement(XLPicturePlacement.FreeFloating)
                           .MoveTo(50, 50);

                Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(400, picture.Width);
                Assert.AreEqual(50, picture.Top);
            }
        }

        [Test]
        public void CanScaleImage()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
                {
                    var pic = ws.AddPicture(resourceStream, "MyPicture")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(50, 50);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(252, pic.Width);
                    Assert.AreEqual(152, pic.Height);

                    pic.ScaleHeight(0.7);
                    pic.ScaleWidth(1.2);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(302, pic.Width);
                    Assert.AreEqual(106, pic.Height);

                    pic.ScaleHeight(0.7);
                    pic.ScaleWidth(1.2);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(362, pic.Width);
                    Assert.AreEqual(74, pic.Height);

                    pic.ScaleHeight(0.8, true);
                    pic.ScaleWidth(1.1, true);

                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(277, pic.Width);
                    Assert.AreEqual(122, pic.Height);
                }
            }
        }

        [Test]
        public void TestDefaultPictureNames()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
                {
                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;
                }

                Assert.AreEqual("Picture 1", ws.Pictures.Skip(0).First().Name);
                Assert.AreEqual("Picture 2", ws.Pictures.Skip(1).First().Name);
                Assert.AreEqual("Picture 4", ws.Pictures.Skip(2).First().Name);
                Assert.AreEqual("Picture 5", ws.Pictures.Skip(3).First().Name);
            }
        }

        [Test]
        public void TestDefaultIds()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
                {
                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png).Name = "Picture 4";
                    stream.Position = 0;

                    ws.AddPicture(stream, XLPictureFormat.Png);
                    stream.Position = 0;
                }

                Assert.AreEqual(1, ws.Pictures.Skip(0).First().Id);
                Assert.AreEqual(2, ws.Pictures.Skip(1).First().Id);
                Assert.AreEqual(3, ws.Pictures.Skip(2).First().Id);
                Assert.AreEqual(4, ws.Pictures.Skip(3).First().Id);
            }
        }

        [Test]
        public void XLMarkerTests()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            XLMarker firstMarker = new XLMarker(ws.Cell(1, 10), new Point(100, 0));

            Assert.AreEqual(10, firstMarker.ColumnNumber);
            Assert.AreEqual(1, firstMarker.RowNumber);
            Assert.AreEqual(100, firstMarker.Offset.X);
            Assert.AreEqual(0, firstMarker.Offset.Y);
        }

        [Test]
        public void XLPictureTests()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
                {
                    var pic = ws.AddPicture(stream, XLPictureFormat.Png, "Image1")
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(220, 155);

                    Assert.AreEqual(XLPicturePlacement.FreeFloating, pic.Placement);
                    Assert.AreEqual("Image1", pic.Name);
                    Assert.AreEqual(XLPictureFormat.Png, pic.Format);
                    Assert.AreEqual(252, pic.OriginalWidth);
                    Assert.AreEqual(152, pic.OriginalHeight);
                    Assert.AreEqual(252, pic.Width);
                    Assert.AreEqual(152, pic.Height);
                    Assert.AreEqual(220, pic.Left);
                    Assert.AreEqual(155, pic.Top);
                }
            }
        }

        [Test]
        public void CanLoadFileWithImagesAndCopyImagesToNewSheet()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                Assert.AreEqual(2, ws.Pictures.Count);

                var copy = ws.CopyTo("NewSheet");
                Assert.AreEqual(2, copy.Pictures.Count);
            }
        }

        [Test]
        public void CanDeletePictures()
        {
            using (var ms = new MemoryStream())
            {
                int originalCount;

                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();
                    originalCount = ws.Pictures.Count;
                    ws.Pictures.Delete(ws.Pictures.First());

                    var pictureName = ws.Pictures.First().Name;
                    ws.Pictures.Delete(pictureName);

                    wb.SaveAs(ms);
                }

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheets.First();
                    Assert.AreEqual(originalCount - 2, ws.Pictures.Count);
                }
            }
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1252")]
        public void Deleted_picture_removes_its_anchor_from_drawing()
        {
            // An anchor that points to a deleted image part makes a corrupt file for Excel.
            using var ms = new MemoryStream();
            int originalCount;
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                originalCount = ws.Pictures.Count;
                ws.Pictures.Delete(ws.Pictures.First());
                wb.SaveAs(ms);
            }

            using (var document = SpreadsheetDocument.Open(ms, false))
            {
                var workbookPart = document.WorkbookPart!;
                var firstSheet = workbookPart.Workbook.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First();
                var drawingsPart = ((WorksheetPart)workbookPart.GetPartById(firstSheet.Id!.Value!)).DrawingsPart!;
                var embeds = drawingsPart.WorksheetDrawing
                    .Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                    .Select(blip => blip.Embed!.Value!)
                    .ToList();
                Assert.AreEqual(originalCount - 1, embeds.Count);
                var partIds = drawingsPart.Parts.Select(part => part.RelationshipId).ToList();
                CollectionAssert.IsSubsetOf(embeds, partIds);
            }

            using (var wb = new XLWorkbook(ms))
            {
                Assert.AreEqual(originalCount - 1, wb.Worksheets.First().Pictures.Count);
            }
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1252")]
        public void Drawing_with_only_shapes_is_kept_on_save()
        {
            // ClosedXML does not load shapes. A shape has no child part, but the drawing part
            // still holds it, so the save must not remove the drawing part.
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet("Sheet1");
                wb.SaveAs(ms);
            }

            using (var document = SpreadsheetDocument.Open(ms, true))
            {
                var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
                var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>("rIdDrawing1");
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(CreateRectangleAnchor(2U));
                var worksheet = worksheetPart.Worksheet;
                worksheet.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = "rIdDrawing1" });
            }

            using (var wb = new XLWorkbook(ms))
            {
                wb.Worksheet("Sheet1").Cell("A1").Value = "Text";
                wb.Save();
            }

            using (var document = SpreadsheetDocument.Open(ms, false))
            {
                var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
                Assert.NotNull(worksheetPart.DrawingsPart);
                var drawingId = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart!);
                Assert.AreEqual(drawingId, worksheetPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>().Single().Id!.Value);
                var shape = worksheetPart.DrawingsPart!.WorksheetDrawing.Descendants<Xdr.Shape>().Single();
                Assert.AreEqual("Rectangle 1", shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value);
            }
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1252")]
        public void Deleted_picture_keeps_shapes_of_drawing()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                using var image = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));
                ws.AddPicture(image, "Picture 1").MoveTo(ws.Cell("E5"));
                wb.SaveAs(ms);
            }

            using (var document = SpreadsheetDocument.Open(ms, true))
            {
                var worksheetDrawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
                worksheetDrawing.AppendChild(CreateRectangleAnchor(100U));
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheet("Sheet1");
                ws.Pictures.Delete("Picture 1");
                wb.Save();
            }

            using (var document = SpreadsheetDocument.Open(ms, false))
            {
                var drawingsPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart;
                Assert.NotNull(drawingsPart);
                CollectionAssert.IsEmpty(drawingsPart!.ImageParts);
                var anchor = drawingsPart.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>().Single();
                Assert.AreEqual("Rectangle 1", anchor.Descendants<Xdr.NonVisualDrawingProperties>().Single().Name!.Value);
            }
        }

        [Test]
        [Issue("ClosedXML/ClosedXML#1252")]
        public void Deleted_picture_removes_each_anchor_of_shared_image()
        {
            // Excel can give one image part to two pictures, for example after a copy of a picture.
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                using var image = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Images\SampleImagePng.png"));
                ws.AddPicture(image, "Picture 1").MoveTo(ws.Cell("B2"));
                wb.SaveAs(ms);
            }

            using (var document = SpreadsheetDocument.Open(ms, true))
            {
                var worksheetDrawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
                var copy = worksheetDrawing.Elements<Xdr.OneCellAnchor>().Single().CloneNode(true);
                var properties = copy.Descendants<Xdr.NonVisualDrawingProperties>().Single();
                properties.Id = 100U;
                properties.Name = "Picture 2";
                worksheetDrawing.AppendChild(copy);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheet("Sheet1");
                Assert.AreEqual(2, ws.Pictures.Count);
                ws.Pictures.Delete("Picture 1");
                wb.Save();
            }

            using (var document = SpreadsheetDocument.Open(ms, false))
            {
                var drawingsPart = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
                var embeds = drawingsPart.WorksheetDrawing
                    .Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                    .Select(blip => blip.Embed!.Value!)
                    .ToList();
                Assert.AreEqual(1, embeds.Count);
                var partIds = drawingsPart.Parts.Select(part => part.RelationshipId).ToList();
                CollectionAssert.IsSubsetOf(embeds, partIds);
            }

            using (var wb = new XLWorkbook(ms))
            {
                Assert.AreEqual("Picture 2", wb.Worksheet("Sheet1").Pictures.Single().Name);
            }
        }

        [Test]
        [Issue("73")]
        public void Loaded_picture_keeps_the_image_data()
        {
            byte[] imageBytes;
            using (var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg"))
            using (var imageStream = new MemoryStream())
            {
                resourceStream.CopyTo(imageStream);
                imageBytes = imageStream.ToArray();
            }

            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet().AddPicture(new MemoryStream(imageBytes), "Picture");
                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var picture = (XLPicture)wb.Worksheets.Single().Pictures.Single();
                Assert.AreEqual(XLPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(0, picture.ImageStream.Position);
                Assert.AreEqual(imageBytes, picture.ImageStream.ToArray());

                // The loaded picture can be saved again.
                using var saved = new MemoryStream();
                wb.SaveAs(saved);
                using var reloaded = new XLWorkbook(saved);
                var reloadedPicture = (XLPicture)reloaded.Worksheets.Single().Pictures.Single();
                Assert.AreEqual(imageBytes, reloadedPicture.ImageStream.ToArray());
            }
        }

        /// <summary>
        /// Add a picture the same way as the load of a workbook adds it.
        /// </summary>
        private static IXLPicture AddLoadedPicture(IXLWorksheet ws, Stream stream, string name, int id)
        {
            var imageStream = new MemoryStream();
            stream.Position = 0;
            stream.CopyTo(imageStream);
            return ((XLPictures)ws.Pictures).AddOwned(imageStream, name, id);
        }

        private static Xdr.TwoCellAnchor CreateRectangleAnchor(uint shapeId)
        {
            return new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(new Xdr.ColumnId("1"), new Xdr.ColumnOffset("0"), new Xdr.RowId("1"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId("3"), new Xdr.ColumnOffset("0"), new Xdr.RowId("4"), new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = shapeId, Name = "Rectangle 1" },
                        new Xdr.NonVisualShapeDrawingProperties()),
                    new Xdr.ShapeProperties(
                        new DocumentFormat.OpenXml.Drawing.PresetGeometry(new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                        { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData());
        }

        [Test]
        public void PictureRenameTests()
        {
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Images3");
                var picture = ws.Pictures.First();
                Assert.AreEqual("Picture 1", picture.Name);

                picture.Name = "picture 1";
                picture.Name = "pICture 1";
                picture.Name = "Picture 1";

                picture = ws.Pictures.Last();
                picture.Name = "new name";

                Assert.Throws<ArgumentException>(() => picture.Name = "Picture 1");
                Assert.Throws<ArgumentException>(() => picture.Name = "picTURE 1");
            }
        }

        [Test]
        public void HandleDuplicatePictureIdsAcrossWorksheets()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                var ws2 = wb.AddWorksheet("Sheet2");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
                {
                    AddLoadedPicture(ws1, stream, "Picture 1", 2);
                    AddLoadedPicture(ws1, stream, "Picture 2", 3);

                    //Internal method - used for loading files
                    var pic = AddLoadedPicture(ws2, stream, "Picture 1", 2)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(220, 155) as XLPicture;

                    var id = pic.Id;

                    pic.Id = id;
                    Assert.AreEqual(id, pic.Id);

                    pic.Id = 3;
                    Assert.AreEqual(3, pic.Id);

                    pic.Id = id;

                    var pic2 = AddLoadedPicture(ws2, stream, "Picture 2", 3)
                        .WithPlacement(XLPicturePlacement.FreeFloating)
                        .MoveTo(440, 300) as XLPicture;
                }
            }
        }

        [Test]
        public void CopyImageSameWorksheet()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");

            IXLPicture original;
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
            {
                original = AddLoadedPicture(ws1, stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(220, 155) as XLPicture;
            }

            var copy = original.Duplicate()
                .MoveTo(300, 200) as XLPicture;

            Assert.AreEqual(2, ws1.Pictures.Count());
            Assert.AreEqual(ws1, copy.Worksheet);
            Assert.AreEqual(original.Format, copy.Format);
            Assert.AreEqual(original.Height, copy.Height);
            Assert.AreEqual(original.Placement, copy.Placement);
            Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
            Assert.AreEqual(original.Width, copy.Width);
            Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");

            Assert.AreEqual(200, copy.Top);
            Assert.AreEqual(300, copy.Left);
            Assert.AreNotEqual(original.Id, copy.Id);
            Assert.AreNotEqual(original.Name, copy.Name);
        }

        [Test]
        public void CopyImageDifferentWorksheets()
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            IXLPicture original;
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
            {
                original = AddLoadedPicture(ws1, stream, "Picture 1", 2)
                    .WithPlacement(XLPicturePlacement.FreeFloating)
                    .MoveTo(220, 155) as XLPicture;
            }
            var ws2 = wb.Worksheets.Add("Sheet2");

            var copy = original.CopyTo(ws2);

            Assert.AreEqual(1, ws1.Pictures.Count());
            Assert.AreEqual(1, ws2.Pictures.Count());

            Assert.AreEqual(ws2, copy.Worksheet);

            Assert.AreEqual(original.Format, copy.Format);
            Assert.AreEqual(original.Height, copy.Height);
            Assert.AreEqual(original.Left, copy.Left);
            Assert.AreEqual(original.Name, copy.Name);
            Assert.AreEqual(original.Placement, copy.Placement);
            Assert.AreEqual(original.Top, copy.Top);
            Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
            Assert.AreEqual(original.Width, copy.Width);
            Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");

            Assert.AreNotEqual(original.Id, copy.Id);
        }

        [Test]
        public void PictureShiftsWhenInsertingRows()
        {
            using (var wb = new XLWorkbook())
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png"))
            {
                var ws = wb.Worksheets.Add("ImageShift");
                var picture = ws.AddPicture(stream, XLPictureFormat.Png, "PngImage")
                    .MoveTo(ws.Cell(5, 2))
                    .WithPlacement(XLPicturePlacement.Move);

                ws.Row(2).InsertRowsBelow(20);

                Assert.AreEqual(25, picture.TopLeftCell.Address.RowNumber);
            }
        }

        [Test]
        public void PictureNotFound()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Picture("dummy"));
                Assert.Throws<ArgumentOutOfRangeException>(() => ws.Pictures.Delete("dummy"));
            }
        }

        [Test]
        public void CanCopyEmfPicture()
        {
            // #1621 - There are 2 Bmp Guids: ImageFormat.Bmp and ImageFormat.MemoryBmp
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Pictures\EmfPicture.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws1 = wb.Worksheets.First();
            var img1 = ws1.Pictures.First();

            var ws2 = wb.AddWorksheet();

            var img2 = img1.CopyTo(ws2);

            Assert.AreEqual(XLPictureFormat.Emf, img2.Format);

            using var ms = new MemoryStream();
            wb.SaveAs(ms);

            ms.Seek(0, SeekOrigin.Begin);

            using var wb2 = new XLWorkbook(ms);
            ws2 = wb2.Worksheet("Sheet2");
            img2 = ws2.Pictures.First();
            Assert.AreEqual(XLPictureFormat.Emf, img2.Format);
        }

        [Test]
        public void KeepOriginalDrawingShapesZOrder()
        {
            // File contains shapes and a picture in a mixed order.
            TestHelper.LoadSaveAndCompare(
                @"Other\Pictures\ImageShapeZOrder-Input.xlsx",
                @"Other\Pictures\ImageShapeZOrder-Output.xlsx");
        }

        [TestCase("@")]
        [TestCase(":")]
        [TestCase("\\")]
        [TestCase("/")]
        [TestCase("?")]
        [TestCase("*")]
        [TestCase("[]")]
        [TestCase(" ")] // Whitespace name is allowed, but can't be empty
        [TestCase("C:\\Images\\pic.jpg")] // Path with multiple forbidden chars
        [TestCase("http://example.com/image.jpg")] // URL with multiple forbidden chars
        [TestCase("Picture@01\\QPosted@")] // A name from a problematic workbook
        public void PictureCanHaveUnusualCharactersInName(string nameWithUnusualCharacter)
        {
            // The name of a picture couldn't contain certain characters in some ancient version of Excel. Verify that
            // it is no longer the case through the whole lifecycle (add picture, change name, save, load).
            AssertPictureNameAllowed(nameWithUnusualCharacter);
        }

        [Test]
        public void PictureNameCanBeLong()
        {
            // Picture name was originally limited to 31 characters. Verify that it is no longer the case.
            var longName = new string('a', 100_000);
            AssertPictureNameAllowed(longName);
        }

        [TestCase(null)]
        [TestCase("")]
        public void PictureNameCantBeNullOrEmpty(string invalidName)
        {
            // Picture name is a required attribute, though Excel generates a name if it isn't specified instead of failing.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            using var imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
            var picture = ws.AddPicture(imageStream);
            Assert.Throws<ArgumentException>(() => picture.Name = invalidName);
        }

        [Test]
        public void Load_gives_a_name_to_a_picture_with_an_empty_name()
        {
            // Excel can save a picture with an empty name. Excel gives such a picture a name when it loads the file.
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                using var imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
                ws.AddPicture(imageStream, "First");
                imageStream.Position = 0;
                ws.AddPicture(imageStream, "Second");
                wb.SaveAs(ms);
            }

            using (var package = SpreadsheetDocument.Open(ms, true))
            {
                var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
                var properties = drawing.Descendants<Xdr.NonVisualDrawingProperties>().ToList();
                Assert.AreEqual(2, properties.Count);
                properties.ForEach(p => p.Name = string.Empty);
                drawing.Save();
            }

            ms.Position = 0;
            using var loaded = new XLWorkbook(ms);
            var names = loaded.Worksheets.Single().Pictures.Select(p => p.Name).ToArray();
            Assert.AreEqual(new[] { "Picture 1", "Picture 2" }, names);
        }

        [Test]
        public void Load_gives_unique_name_to_unnamed_picture_when_later_picture_has_generated_name()
        {
            // The generated name of the unnamed picture is "Picture 1". The next picture has this name in the file.
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                using var imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
                ws.AddPicture(imageStream, "First");
                imageStream.Position = 0;
                ws.AddPicture(imageStream, "Picture 1");
                wb.SaveAs(ms);
            }

            using (var package = SpreadsheetDocument.Open(ms, true))
            {
                var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
                var first = drawing.Descendants<Xdr.NonVisualDrawingProperties>().Single(p => p.Name == "First");
                first.Name = string.Empty;
                drawing.Save();
            }

            ms.Position = 0;
            using var loaded = new XLWorkbook(ms);
            var names = loaded.Worksheets.Single().Pictures.Select(p => p.Name).ToArray();
            Assert.AreEqual(new[] { "Picture 2", "Picture 1" }, names);
        }

        [Test]
        public void Generated_picture_name_differs_from_names_with_other_case()
        {
            // Picture names are case-insensitive, so "Picture 3" is not available when "picture 3" exists.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            using var imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
            ws.AddPicture(imageStream, "First");
            imageStream.Position = 0;
            ws.AddPicture(imageStream, "picture 3");
            imageStream.Position = 0;

            var picture = ws.AddPicture(imageStream);

            Assert.AreEqual("Picture 4", picture.Name);
        }

        private static void AssertPictureNameAllowed(string testedName)
        {
            TestHelper.CreateSaveLoadAssert(
                wb =>
                {
                    using var imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ClosedXML.Tests.Resource.Images.ImageHandling.png");
                    var ws1 = wb.AddWorksheet("AddPicture");
                    ws1.AddPicture(imageStream, testedName);

                    var ws2 = wb.AddWorksheet("Setter");
                    var picture = ws2.AddPicture(imageStream);
                    picture.Name = testedName;
                }, wb =>
                {
                    Assert.AreEqual(testedName, wb.Worksheet("AddPicture").Pictures.Single().Name);
                    Assert.AreEqual(testedName, wb.Worksheet("Setter").Pictures.Single().Name);
                });
        }

        [Test]
        [Issue("64")]
        public void Picture_without_anchors_has_no_anchor_cells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            using var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg");

            var picture = ws.AddPicture(resourceStream);

            Assert.IsNull(picture.TopLeftCell);
            Assert.IsNull(picture.BottomRightCell);
            Assert.Throws<InvalidOperationException>(() => picture.GetOffset(XLMarkerPosition.TopLeft));
        }

        [Test]
        [Issue("64")]
        public void Free_floating_picture_has_no_bottom_right_anchor()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            using var resourceStream = Assembly.GetAssembly(typeof(ClosedXML.Examples.BasicTable)).GetManifestResourceStream("ClosedXML.Examples.Resources.SampleImage.jpg");

            var picture = ws.AddPicture(resourceStream).MoveTo(50, 50);

            Assert.AreEqual("A1", picture.TopLeftCell.Address.ToString());
            Assert.IsNull(picture.BottomRightCell);
            Assert.Throws<InvalidOperationException>(() => picture.GetOffset(XLMarkerPosition.BottomRight));
        }
    }
}
