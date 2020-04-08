using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using System;
using System.Drawing;
using System.Linq;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Shapes
{
    internal class WorkingWithShapes : TestDataHelper
    {
        public static void Run()
        {
            SetShapeLayoutInCell();
            SetAspectRatioLocked();
            InsertShapeUsingDocumentBuilder();
            AddCornersSnipped();
            GetActualShapeBoundsPoints();
            SpecifyVerticalAnchor();
            DetectSmartArtShape();
            InsertOleObjectAsIcon();
        }

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

        public static void SetShapeLayoutInCell()
        {
            //ExStart:SetShapeLayoutInCell
            Document doc = new Document(ShapesDir + "LayoutInCell.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

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

        public static void SpecifyVerticalAnchor()
        {
            //ExStart:SpecifyVerticalAnchor
            Document doc = new Document(ShapesDir + "VerticalAnchor.docx");
            
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            if (shapes[0] is Shape textBoxShape)
            {
                textBoxShape.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            }

            doc.Save(ArtifactsDir + "VerticalAnchor.docx");
            //ExEnd:SpecifyVerticalAnchor
        }

        public static void DetectSmartArtShape()
        {
            //ExStart:DetectSmartArtShape
            Document doc = new Document(ShapesDir + "input.docx");

            int count = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Count(shape => shape.HasSmartArt);

            Console.WriteLine("The document has {0} shapes with SmartArt.", count);
            //ExEnd:DetectSmartArtShape
        }

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