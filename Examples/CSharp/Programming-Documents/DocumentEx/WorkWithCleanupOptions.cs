using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WorkWithCleanupOptions : TestDataHelper
    {
        [Test]
        public static void CleanupUnusedStylesandLists()
        {
            // ExStart:CleanupUnusedStylesandLists
            Document doc = new Document(DocumentDir + "Document.doc");

            CleanupOptions cleanupOptions = new CleanupOptions();
            cleanupOptions.UnusedLists = false;
            cleanupOptions.UnusedStyles = true;

            // Cleans unused styles and lists from the document depending on given CleanupOptions. 
            doc.Cleanup(cleanupOptions);

            doc.Save(ArtifactsDir + "Document.CleanupUnusedStylesandLists.docx");
            // ExEnd:CleanupUnusedStylesandLists
        }

        [Test]
        public static void CleanupDuplicateStyle()
        {
            // ExStart:CleanupDuplicateStyle
            Document doc = new Document(DocumentDir + "Document.doc");

            CleanupOptions options = new CleanupOptions();
            options.DuplicateStyle = true;

            // Cleans duplicate styles from the document. 
            doc.Cleanup(options);

            doc.Save(ArtifactsDir + "Document.CleanupDuplicateStyle_out.docx");
            // ExEnd:CleanupDuplicateStyle
        }
    }
}
