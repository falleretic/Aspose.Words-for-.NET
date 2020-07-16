using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class ShowGrammaticalAndSpellingErrors : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // ExStart: ShowGrammaticalAndSpellingErrors
            Document doc = new Document(DocumentDir + "Document.doc");

            doc.ShowGrammaticalErrors = true;
            doc.ShowSpellingErrors = true;

            doc.Save(ArtifactsDir + "Document.ShowErrorsInDocument.docx");
            // ExEnd: ShowGrammaticalAndSpellingErrors
        }
    }
}
