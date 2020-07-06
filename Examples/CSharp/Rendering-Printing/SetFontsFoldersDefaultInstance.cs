using Aspose.Words.Fonts;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class SetFontsFoldersDefaultInstance : TestDataHelper
    {
        public static void Run()
        {
            // ExStart:SetFontsFoldersDefaultInstance
            FontSettings.DefaultInstance.SetFontsFolder("C:\\MyFonts\\", true);
            // ExEnd:SetFontsFoldersDefaultInstance           

            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");
            doc.Save(ArtifactsDir + "Rendering.SetFontsFolders.pdf");
        }
    }
}
