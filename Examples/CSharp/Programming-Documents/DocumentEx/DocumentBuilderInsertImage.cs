using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class DocumentBuilderInsertImage : TestDataHelper
    {
        [Test]
        public static void InsertInlineImage()
        {
            //ExStart:DocumentBuilderInsertInlineImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertImage(DocumentDir + "Watermark.png");
            
            doc.Save(ArtifactsDir + "DocumentBuilderInsertInlineImage.doc");
            //ExEnd:DocumentBuilderInsertInlineImage
        }

        [Test]
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