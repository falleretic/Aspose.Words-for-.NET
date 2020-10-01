using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class HtmlFixedSaveOptionsEx : TestDataHelper
    {
        [Test]
        public static void UseFontFromTargetMachine()
        {
            //ExStart:UseFontFromTargetMachine
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            saveOptions.UseTargetMachineFonts = true;

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UseFontFromTargetMachine.html", saveOptions);
            //ExEnd:UseFontFromTargetMachine
        }

        [Test]
        public static void WriteAllCssRulesInSingleFile()
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