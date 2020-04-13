using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToPCL : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ConvertDocumentToPCL
            Document doc = new Document(LoadingSavingDir + "Test File (docx).docx");

            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Pcl;
            saveOptions.RasterizeTransformedElements = false;

            // Export the document as an PCL file
            doc.Save(ArtifactsDir + "ConvertDocumentToPCL.pcl", saveOptions);
            //ExEnd:ConvertDocumentToPCL
        }
    }
}