using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class WorkingWithPclSaveOptions : TestDataHelper
    {
        [Test, Description("Shows how not to rasterize transformed elements.")]
        public static void RasterizeTransformedElements()
        {
            //ExStart:RasterizeTransformedElements
            Document doc = new Document(MyDir + "Rendering.docx");

            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Pcl;
            saveOptions.RasterizeTransformedElements = false;

            doc.Save(ArtifactsDir + "PclSaveOptions.RasterizeTransformedElements.pcl", saveOptions);
            //ExEnd:RasterizeTransformedElements
        }
    }
}