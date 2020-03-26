using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertImage : TestDataHelper
    {
        public static void Run()
        {
            InsertInlineImage();
            InsertFloatingImage();
        }

        public static void InsertInlineImage()
        {
            //ExStart:DocumentBuilderInsertInlineImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertImage(DocumentDir + "Watermark.png");
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertInlineImage.doc");
            //ExEnd:DocumentBuilderInsertInlineImage
        }

        public static void InsertFloatingImage()
        {
            //ExStart:DocumentBuilderInsertFloatingImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(DocumentDir + "Watermark.png",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertFloatingImage.doc");
            //ExEnd:DocumentBuilderInsertFloatingImage
        }
    }
}