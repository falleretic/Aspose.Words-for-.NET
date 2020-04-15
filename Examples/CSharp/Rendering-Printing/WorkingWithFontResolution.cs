using Aspose.Words.Fonts;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    class WorkingWithFontResolution : TestDataHelper
    {
        [Test]
        public static void FontSettingsWithLoadOptions()
        {
            //ExStart:FontSettingsWithLoadOptions
            FontSettings fontSettings = new FontSettings();

            TableSubstitutionRule substitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;
            // If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS"
            substitutionRule.AddSubstitutes("UnknownFont1", new string[] { "Comic Sans MS" });
            
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            
            Document doc = new Document(ArtifactsDir + "myfile.html", loadOptions);
            //ExEnd:FontSettingsWithLoadOptions
        }

        [Test]
        public static void SetFontsFolder()
        {
            //ExStart:SetFontsFolder
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolder(MailMergeDir + "Fonts", false);
            
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            
            Document doc = new Document(MailMergeDir + "myfile.html", loadOptions);
            //ExEnd:SetFontsFolder
        }
    }
}