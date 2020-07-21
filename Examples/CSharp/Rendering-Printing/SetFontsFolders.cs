using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class SetFontsFolders : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // ExStart:SetFontsFolders
            FontSettings.DefaultInstance.SetFontsSources(new FontSourceBase[]
            {
                new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true)
            });

            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "Rendering.SetFontsFolders.pdf");
            // ExEnd:SetFontsFolders           
        }
    }
}
