using Aspose.Words.Fonts;
using System.Collections;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetFontsFoldersSystemAndCustomFolder : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SetFontsFoldersSystemAndCustomFolder
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");
            
            FontSettings fontSettings = new FontSettings();
            // Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines
            // We add this array to a new ArrayList to make adding or removing font entries much easier
            ArrayList fontSources = new ArrayList(fontSettings.GetFontsSources());
            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
            FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);
            // Add the custom folder which contains our fonts to the list of existing font sources
            fontSources.Add(folderFontSource);
            // Convert the ArrayList of source back into a primitive array of FontSource objects
            FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources.ToArray(typeof(FontSourceBase));
            // Apply the new set of font sources to use
            fontSettings.SetFontsSources(updatedFontSources);
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "SetFontsFoldersSystemAndCustomFolder.pdf");
            //ExEnd:SetFontsFoldersSystemAndCustomFolder
        }
    }
}