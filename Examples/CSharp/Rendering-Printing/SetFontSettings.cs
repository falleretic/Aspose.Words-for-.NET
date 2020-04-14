using Aspose.Words.Fonts;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SetFontSettings : TestDataHelper
    {
        public static void Run()
        {
            EnableDisableFontSubstitution();
            SetFontFallbackSettings();
            SetPredefinedFontFallbackSettings();
        }

        public static void EnableDisableFontSubstitution()
        {
            //ExStart:EnableDisableFontSubstitution
            Document doc = new Document(MailMergeDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "EnableDisableFontSubstitution.pdf");
            //ExEnd:EnableDisableFontSubstitution
        }

        public static void SetFontFallbackSettings()
        {
            //ExStart:SetFontFallbackSettings
            Document doc = new Document(MailMergeDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.Load(MailMergeDir + "Fallback.xml");
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "SetFontFallbackSettings.pdf");
            //ExEnd:SetFontFallbackSettings
        }

        public static void SetPredefinedFontFallbackSettings()
        {
            //ExStart:SetPredefinedFontFallbackSettings
            Document doc = new Document(MailMergeDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.LoadNotoFallbackSettings();
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "SetPredefinedFontFallbackSettings.pdf");
            //ExEnd:SetPredefinedFontFallbackSettings
        }
    }
}