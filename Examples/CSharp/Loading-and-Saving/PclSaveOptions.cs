using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ConvertDocumentToPCL : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ConvertDocumentToPCL
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");

            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Pcl;
            saveOptions.RasterizeTransformedElements = false;

            // Export the document as an PCL file
            doc.Save(ArtifactsDir + "ConvertDocumentToPCL.pcl", saveOptions);
            //ExEnd:ConvertDocumentToPCL
        }
    }
}