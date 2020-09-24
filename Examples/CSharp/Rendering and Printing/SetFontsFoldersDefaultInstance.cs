using Aspose.Words.Fonts;

namespace Aspose.Words.Examples.CSharp
{
    class SetFontsFoldersDefaultInstance : TestDataHelper
    {
        public static void Run()
        {
            // ExStart:SetFontsFoldersDefaultInstance
            FontSettings.DefaultInstance.SetFontsFolder("C:\\MyFonts\\", true);
            // ExEnd:SetFontsFoldersDefaultInstance           

            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "Rendering.SetFontsFolders.pdf");
        }
    }
}
