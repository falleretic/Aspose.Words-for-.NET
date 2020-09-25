using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class SpecificDocumentOptions : TestDataHelper
    {
        [Test]
        public static void OptimizeFor()
        {
            //ExStart:OptimizeFor
            Document doc = new Document(DocumentDir + "Document.docx");
            doc.CompatibilityOptions.OptimizeFor(Settings.MsWordVersion.Word2016);

            doc.Save(ArtifactsDir + "TestFile.docx");
            //ExEnd:OptimizeFor
        }

        [Test]
        public static void ShowGrammaticalAndSpellingErrors()
        {
            // ExStart: ShowGrammaticalAndSpellingErrors
            Document doc = new Document(DocumentDir + "Document.docx");

            doc.ShowGrammaticalErrors = true;
            doc.ShowSpellingErrors = true;

            doc.Save(ArtifactsDir + "Document.ShowErrorsInDocument.docx");
            // ExEnd: ShowGrammaticalAndSpellingErrors
        }

        [Test]
        public static void CleanupUnusedStylesandLists()
        {
            // ExStart:CleanupUnusedStylesandLists
            Document doc = new Document(DocumentDir + "Document.docx");

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
            Document doc = new Document(DocumentDir + "Document.docx");

            CleanupOptions options = new CleanupOptions();
            options.DuplicateStyle = true;

            // Cleans duplicate styles from the document. 
            doc.Cleanup(options);

            doc.Save(ArtifactsDir + "Document.CleanupDuplicateStyle_out.docx");
            // ExEnd:CleanupDuplicateStyle
        }
    }
}