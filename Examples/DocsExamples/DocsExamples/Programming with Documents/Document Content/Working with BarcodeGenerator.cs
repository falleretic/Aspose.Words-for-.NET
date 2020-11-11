using Aspose.Words;
using DocsExamples.Programming_with_Documents.Document_Content.Helpers;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    internal class WorkingWithBarcodeGenerator : DocsExamplesBase
    {
        [Test]
        public static void GenerateACustomBarCodeImage()
        {
            //ExStart:GenerateACustomBarCodeImage
            Document doc = new Document(MyDir + "Field sample - BARCODE.docx");

            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
            
            doc.Save(ArtifactsDir + "WorkingWithBarcodeGenerator.GenerateACustomBarCodeImage.pdf");
            //ExEnd:GenerateACustomBarCodeImage
        }
    }
}