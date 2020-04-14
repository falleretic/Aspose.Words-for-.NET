﻿using Aspose.Words.Fonts;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetTrueTypeFontsFolder : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:SetTrueTypeFontsFolder
            Document doc = new Document(MailMergeDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
            // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
            // FontSettings.SetFontSources instead
            fontSettings.SetFontsFolder(@"C:\MyFonts\", false);
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "SetTrueTypeFontsFolder.pdf");
            //ExEnd:SetTrueTypeFontsFolder
        }
    }
}