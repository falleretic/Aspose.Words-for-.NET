namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadAndSaveHtmlFormFieldAsContentControlInDocx : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:LoadAndSaveHtmlFormFieldasContentControlinDOCX
            HtmlLoadOptions lo = new HtmlLoadOptions();
            lo.PreferredControlType = HtmlControlType.StructuredDocumentTag;

            Document doc = new Document(LoadingSavingDir + "input.html", lo);
            doc.Save(ArtifactsDir + "LoadAndSaveHtmlFormFieldAsContentControlInDocx.docx", SaveFormat.Docx);
            //ExEnd:LoadAndSaveHtmlFormFieldasContentControlinDOCX
        }
    }
}