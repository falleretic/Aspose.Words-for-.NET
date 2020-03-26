namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CleanUnusedStylesAndLists : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:CleansUnusedStylesandLists
            Document doc = new Document(DocumentDir + "Document.doc");

            CleanupOptions cleanupOptions = new CleanupOptions();
            cleanupOptions.UnusedLists = false;
            cleanupOptions.UnusedStyles = true;

            // Clean unused styles and lists from the document depending on given CleanupOptions
            doc.Cleanup(cleanupOptions);

            doc.Save(ArtifactsDir + "Document.Cleanup.docx");
            //ExEnd:CleansUnusedStylesandLists
        }
    }
}