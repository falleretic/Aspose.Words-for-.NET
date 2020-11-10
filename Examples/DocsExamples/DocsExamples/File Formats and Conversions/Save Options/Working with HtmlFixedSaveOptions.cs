using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithHtmlFixedSaveOptions : DocsExamplesBase
    {
        [Test]
        public void UseFontFromTargetMachine()
        {
            //ExStart:UseFontFromTargetMachine
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions { UseTargetMachineFonts = true };

            doc.Save(ArtifactsDir + "WorkingWithHtmlFixedSaveOptions.UseFontFromTargetMachine.html", htmlFixedSaveOptions);
            //ExEnd:UseFontFromTargetMachine
        }

        [Test]
        public void WriteAllCssRulesInSingleFile()
        {
            //ExStart:WriteAllCssRulesInSingleFile
            Document doc = new Document(MyDir + "Document.docx");

            // Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
            // All CSS rules are written into single file "styles.css.
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions { SaveFontFaceCssSeparately = false };
            
            doc.Save(ArtifactsDir + "WorkingWithHtmlFixedSaveOptions.WriteAllCssRulesInSingleFile.html", htmlFixedSaveOptions);
            //ExEnd:WriteAllCssRulesInSingleFile
        }
    }
}