using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class SetFontSettings : TestDataHelper
    {
        [Test]
        public static void EnableDisableFontSubstitution()
        {
            //ExStart:EnableDisableFontSubstitution
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "EnableDisableFontSubstitution.pdf");
            //ExEnd:EnableDisableFontSubstitution
        }

        [Test]
        public static void SetFontFallbackSettings()
        {
            //ExStart:SetFontFallbackSettings
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.Load(RenderingPrintingDir + "Fallback.xml");
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "SetFontFallbackSettings.pdf");
            //ExEnd:SetFontFallbackSettings
        }

        [Test]
        public static void SetPredefinedFontFallbackSettings()
        {
            //ExStart:SetPredefinedFontFallbackSettings
            Document doc = new Document(RenderingPrintingDir + "Rendering.doc");

            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.LoadNotoFallbackSettings();
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "SetPredefinedFontFallbackSettings.pdf");
            //ExEnd:SetPredefinedFontFallbackSettings
        }
    }
}