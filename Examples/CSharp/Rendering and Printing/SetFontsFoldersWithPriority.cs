using Aspose.Words.Fonts;

namespace Aspose.Words.Examples.CSharp
{
    class SetFontsFoldersWithPriority : TestDataHelper
    {
        public static void Run()
        {
            // ExStart:SetFontsFoldersWithPriority
            FontSettings.DefaultInstance.SetFontsSources(new FontSourceBase[]
            {
                new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true,1)
            });
            // ExEnd:SetFontsFoldersWithPriority           

            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "Rendering.SetFontsFolders.pdf");
        }
    }
}
