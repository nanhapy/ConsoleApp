using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SearchAndReplace("contract.dotx", new Parent
            {
                Name = "Papa",
                Age = 29,
                Avatar = "minjie-seal.png",
                Children = new List<Child>() { new Child {
                    Name = "Yuanyuan",
                    Age = 1
                },
                new Child {
                    Name = "Jinjin",
                    Age = 2
                }}
            });
        }

        static void ZhongGu()
        {
            //var client = new HttpClient();
            //var postData = new { waybillNum = "ZGZ1727STJQZ147", boxNum = "ZGXU2014764" };
            //string json = Newtonsoft.Json.JsonConvert.SerializeObject(postData);
            //HttpContent contentPost = new StringContent(json, Encoding.UTF8, "application/json");
            //HttpResponseMessage response = client.PostAsync("http://dingcang.xlhy.cn/waybill/web/dynamic.json", contentPost).Result;
            //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
        }

        public static void SearchAndReplace(string document, Parent parent)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.CreateFromTemplate(document))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                docText = docText.Replace("{order.orderId}", "1111222222222222222")
                .Replace("{order.customerName}", "盛安德盛安德盛安德盛安德盛安德")
                .Replace("{order.customerAddress}", "盛安德盛安德盛安德盛安德盛安德Address")
                .Replace("{order.customerContact}", "小张张")
                .Replace("{order.customerFaxNumber}", "34565675644545")
                .Replace("{order.selfName}", "盛安德盛安德盛安德盛安德盛安德")
                .Replace("{order.signDate}", "1990-10-12")
                .Replace("{order.endWharf}", "曹妃甸")
                .Replace("{order.startWharf}", "天津港");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }

                Body bod = wordDoc.MainDocumentPart.Document.Body;
                Table table = bod.Descendants<Table>().First(t => t.InnerText.Contains("船名/航次"));
                var rows = table.Descendants<TableRow>();
                var templateRow = table.Descendants<TableRow>().ElementAt(1);
                int len = 1;
                while (len <= 10)
                {
                    var newRow = templateRow.CloneNode(true) as TableRow;
                    var cells = newRow.Descendants<TableCell>();
                    ReplaceTableCellText(cells, "{item.voyage}", "WAN MU CHUN 10/1801");
                    ReplaceTableCellText(cells, "{item.cargo}", "角钢");
                    ReplaceTableCellText(cells, "{item.value}", "300000.00");
                    ReplaceTableCellText(cells, "{item.amount}", "40.00");
                    ReplaceTableCellText(cells, "{item.weight}", "200000.00");
                    ReplaceTableCellText(cells, "{item.bgfUnit}", "100000.00 元/吨");
                    ReplaceTableCellText(cells, "{item.receivable}", "-1474836480.00 元");

                    templateRow.InsertAfterSelf(newRow);
                    len++;
                }
                templateRow.Remove();

                ReplaceBookmarksWithImage(wordDoc, "seal", "minjie-seal.png");

                wordDoc.SaveAs("demo.docx");
            }
        }

        static void ReplaceTableCellText(IEnumerable<TableCell> cells, string Key, string text)
        {
            try
            {
                var cellCargo = cells.First(c => c.InnerText == Key);
                cellCargo.RemoveAllChildren<Paragraph>();
                cellCargo.Append(new Paragraph(new Run(new Text(text))));
            }
            catch (InvalidOperationException ioe)
            {

            }
        }

        private static void ReplaceBookmarksWithImage(WordprocessingDocument doc, string bookmarkStartName, string imageFilename)
        {
            // Read all bookmarks from the word doc
            foreach (BookmarkStart bookmarkStart in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                if (bookmarkStart.Name == bookmarkStartName)
                {
                    // insert the image
                    InsertImageIntoBookmark(doc, bookmarkStart, imageFilename);

                    // remove the bookmark
                    bookmarkStart.Remove();
                }
            }
        }

        public static void InsertImageIntoBookmark(WordprocessingDocument doc, BookmarkStart bookmarkStart, string imageFilename)
        {
            // Remove anything present inside the bookmark
            OpenXmlElement elem = bookmarkStart.NextSibling();
            while (elem != null && !(elem is BookmarkEnd))
            {
                OpenXmlElement nextElem = elem.NextSibling();
                elem.Remove();
                elem = nextElem;
            }

            // Create an imagepart
            var imagePart = AddImagePart(doc.MainDocumentPart, imageFilename);

            // insert the image part after the bookmark start
            AddImageToBody(doc.MainDocumentPart.GetIdOfPart(imagePart), bookmarkStart);
        }

        public static ImagePart AddImagePart(MainDocumentPart mainPart, string imageFilename)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(imageFilename, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            return imagePart;
        }

        private static void AddImageToBody(string relationshipId, BookmarkStart bookmarkStart)
        {
            // Define the reference of the image.
            Drawing drawing1 = new Drawing();

            DW.Anchor anchor1 = new DW.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251660288U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true };
            DW.SimplePosition simplePosition1 = new DW.SimplePosition() { X = 0L, Y = 0L };

            DW.HorizontalPosition horizontalPosition1 = new DW.HorizontalPosition() { RelativeFrom = DW.HorizontalRelativePositionValues.Column };
            DW.PositionOffset positionOffset1 = new DW.PositionOffset();
            positionOffset1.Text = "907415";

            horizontalPosition1.Append(positionOffset1);

            DW.VerticalPosition verticalPosition1 = new DW.VerticalPosition() { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph };
            DW.PositionOffset positionOffset2 = new DW.PositionOffset();
            positionOffset2.Text = "-2540";

            verticalPosition1.Append(positionOffset2);
            DW.Extent extent1 = new DW.Extent() { Cx = 989965L, Cy = 989965L };
            DW.EffectExtent effectExtent1 = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 635L, BottomEdge = 635L };
            DW.WrapNone wrapNone1 = new DW.WrapNone();
            DW.DocProperties docProperties1 = new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" };

            DW.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new DW.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            PIC.Picture picture1 = new PIC.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            PIC.NonVisualPictureProperties nonVisualPictureProperties1 = new PIC.NonVisualPictureProperties();
            PIC.NonVisualDrawingProperties nonVisualDrawingProperties1 = new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "New Bitmap Image.jpg" };
            PIC.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new PIC.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            PIC.BlipFill blipFill1 = new PIC.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            PIC.ShapeProperties shapeProperties1 = new PIC.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 989965L, Cy = 989965L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);

            drawing1.Append(anchor1);


            // add the image element to body, the element should be in a Run.
            bookmarkStart.Parent.InsertAfter<Run>(new Run(drawing1), bookmarkStart);
        }
    }

    public class Queryable<T>
    {
    }

    public class Parent
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public List<Child> Children { get; set; }
        public string Avatar { get; set; }
    }

    public class Child
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class ChildB : Parent { }

    public class WordParameter
    {
        public string Name { get; set; }
        public string Text { get; set; }
        public FileInfo Image { get; set; }
    }
}
