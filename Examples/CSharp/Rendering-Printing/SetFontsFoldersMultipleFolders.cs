using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetFontsFoldersMultipleFolders : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:SetFontsFoldersMultipleFolders
            Document doc = new Document(MailMergeDir + "Rendering.doc");
            
            FontSettings fontSettings = new FontSettings();
            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
            // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
            // FontSettings.SetFontSources instead
            fontSettings.SetFontsFolders(new[] { @"C:\MyFonts\", @"D:\Misc\Fonts\" }, true);
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "SetFontsFoldersMultipleFolders.pdf");
            //ExEnd:SetFontsFoldersMultipleFolders           
        }
    }
}