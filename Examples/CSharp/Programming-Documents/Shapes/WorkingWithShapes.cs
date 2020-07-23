using System;
using System.Drawing;
using System.Linq;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Shapes
{
    internal class WorkingWithShapes : TestDataHelper
    {
        [Test]
        public static void InsertShapeUsingDocumentBuilder()
        {
            //ExStart:InsertShapeUsingDocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Free-floating shape insertion
            Shape shape = builder.InsertShape(ShapeType.TextBox, RelativeHorizontalPosition.Page, 100,
                RelativeVerticalPosition.Page, 100, 50, 50, WrapType.None);
            shape.Rotation = 30.0;

            builder.Writeln();

            // Inline shape insertion
            shape = builder.InsertShape(ShapeType.TextBox, 50, 50);
            shape.Rotation = 30.0;

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            // "Strict" or "Transitional" compliance allows to save shape as DML
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

            doc.Save(ArtifactsDir + "Shape_InsertShapeUsingDocumentBuilder.docx", saveOptions);
            //ExEnd:InsertShapeUsingDocumentBuilder
        }

        [Test]
        public static void SetAspectRatioLocked()
        {
            //ExStart:SetAspectRatioLocked
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertImage(ShapesDir + "Test.png");
            shape.AspectRatioLocked = false;

            doc.Save(ArtifactsDir + "Shape_AspectRatioLocked.doc");
            //ExEnd:SetAspectRatioLocked
        }

        [Test]
        public static void SetShapeLayoutInCell()
        {
            //ExStart:SetShapeLayoutInCell
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            for (int i = 0; i < 31; i++)
            {
                if (i != 0 && i % 7 == 0) builder.EndRow();
                builder.InsertCell();
                builder.Write("Cell contents");
            }

            builder.EndTable();

            Shape watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            // Display the shape outside of table cell if it will be placed into a cell
            watermark.IsLayoutInCell = true;

            watermark.Width = 300;
            watermark.Height = 70;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.VerticalAlignment = VerticalAlignment.Center;

            watermark.Rotation = -40;
            watermark.Fill.Color = Color.Gray;
            watermark.StrokeColor = Color.Gray;

            watermark.TextPath.Text = "watermarkText";
            watermark.TextPath.FontFamily = "Arial";

            watermark.Name = string.Format("WaterMark_{0}", Guid.NewGuid());
            watermark.WrapType = WrapType.None;

            Run run = doc.GetChildNodes(NodeType.Run, true)[doc.GetChildNodes(NodeType.Run, true).Count - 1] as Run;

            builder.MoveTo(run);
            builder.InsertNode(watermark);
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

            doc.Save(ArtifactsDir + "Shape_IsLayoutInCell.docx");
            //ExEnd:SetShapeLayoutInCell
        }

        [Test]
        public static void AddCornersSnipped()
        {
            //ExStart:AddCornersSnipped
            Document doc = new Document();
            
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertShape(ShapeType.TopCornersSnipped, 50, 50);

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
            
            doc.Save(ArtifactsDir + "AddCornersSnipped.docx", saveOptions);
            //ExEnd:AddCornersSnipped
        }

        [Test]
        public static void GetActualShapeBoundsPoints()
        {
            //ExStart:GetActualShapeBoundsPoints
            Document doc = new Document();
            
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape shape = builder.InsertImage(ShapesDir + "Test.png");
            shape.AspectRatioLocked = false;

            Console.Write("\nGets the actual bounds of the shape in points: ");
            Console.WriteLine(shape.GetShapeRenderer().BoundsInPoints);
            //ExEnd:GetActualShapeBoundsPoints
        }

        [Test]
        public static void SpecifyVerticalAnchor()
        {
            //ExStart:SpecifyVerticalAnchor
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
            textBox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Textbox contents");

            doc.Save(ArtifactsDir + "VerticalAnchor.docx");
            //ExEnd:SpecifyVerticalAnchor
        }

        [Test]
        public static void DetectSmartArtShape()
        {
            //ExStart:DetectSmartArtShape
            Document doc = new Document(ShapesDir + "SmartArt.docx");

            int count = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Count(shape => shape.HasSmartArt);

            Console.WriteLine("The document has {0} shapes with SmartArt.", count);
            //ExEnd:DetectSmartArtShape
        }

        [Test]
        public static void InsertOleObjectAsIcon()
        {
            //ExStart:InsertOLEObjectAsIcon
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertOleObjectAsIcon(ShapesDir + "embedded.xlsx", false, ShapesDir + "icon.ico",
                "My embedded file");

            doc.Save(ArtifactsDir + "EmbeddeWithIcon.docx");

            Console.WriteLine("The document has been saved with OLE Object as an Icon.");
            //ExEnd:InsertOLEObjectAsIcon
        }
    }
}