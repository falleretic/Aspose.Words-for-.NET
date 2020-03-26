using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class AddGroupShapeToDocument : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:AddGroupShapeToDocument
            Document doc = new Document();
            doc.EnsureMinimum();
            
            GroupShape gs = new GroupShape(doc);
            Shape shape = new Shape(doc, ShapeType.AccentBorderCallout1);
            shape.Width = 100;
            shape.Height = 100;
            gs.AppendChild(shape);

            Shape shape1 = new Shape(doc, ShapeType.ActionButtonBeginning);
            shape1.Left = 100;
            shape1.Width = 100;
            shape1.Height = 200;
            gs.AppendChild(shape1);
            
            gs.Width = 200;
            gs.Height = 200;
            gs.CoordSize = new System.Drawing.Size(200, 200);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertNode(gs);

            doc.Save(ArtifactsDir + "groupshape-doc.doc");
            //ExEnd:AddGroupShapeToDocument
        }
    }
}