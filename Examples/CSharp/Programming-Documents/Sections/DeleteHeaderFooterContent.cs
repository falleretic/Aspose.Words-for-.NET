using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Sections
{
    class DeleteHeaderFooterContent : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:DeleteHeaderFooterContent
            Document doc = new Document(SectionsDir + "Document.docx");
            
            Section section = doc.Sections[0];
            section.ClearHeadersFooters();
            //ExEnd:DeleteHeaderFooterContent
        }
    }
}