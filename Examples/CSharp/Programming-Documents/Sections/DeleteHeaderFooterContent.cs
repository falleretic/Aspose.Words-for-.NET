namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class DeleteHeaderFooterContent : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:DeleteHeaderFooterContent
            Document doc = new Document(SectionsDir + "Document.doc");
            
            Section section = doc.Sections[0];
            section.ClearHeadersFooters();
            //ExEnd:DeleteHeaderFooterContent
        }
    }
}