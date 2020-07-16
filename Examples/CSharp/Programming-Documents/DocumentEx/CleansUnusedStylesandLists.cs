using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class CleanUnusedStylesAndLists : TestDataHelper
    {
        [Test]
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