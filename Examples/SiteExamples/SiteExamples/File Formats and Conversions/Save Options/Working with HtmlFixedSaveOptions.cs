using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Save_Options
{
    internal class HtmlFixedSaveOptionsEx : SiteExamplesBase
    {
        [Test, Description("Shows how to use fonts from the target machine.")]
        public void UseFontFromTargetMachine()
        {
            //ExStart:UseFontFromTargetMachine
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            saveOptions.UseTargetMachineFonts = true;

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UseFontFromTargetMachine.html", saveOptions);
            //ExEnd:UseFontFromTargetMachine
        }

        [Test, Description("Shows how to create separate fontFaces.css.")]
        public void WriteAllCssRulesInSingleFile()
        {
            //ExStart:WriteAllCssRulesInSingleFile
            Document doc = new Document(MyDir + "Document.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            // Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
            // All CSS rules are written into single file "styles.css.
            saveOptions.SaveFontFaceCssSeparately = false;

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.WriteAllCssRulesInSingleFile.html", saveOptions);
            //ExEnd:WriteAllCssRulesInSingleFile
        }
    }
}