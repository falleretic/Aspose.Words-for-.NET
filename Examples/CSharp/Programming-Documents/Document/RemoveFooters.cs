using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class RemoveFooters : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:RemoveFooters
            Document doc = new Document(DocumentDir + "HeaderFooter.RemoveFooters.doc");

            foreach (Section section in doc)
            {
                // Up to three different footers are possible in a section (for first, even and odd pages)
                // We check and delete all of them
                HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
                footer?.Remove();

                // Primary footer is the footer used for odd pages
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                footer?.Remove();

                footer = section.HeadersFooters[HeaderFooterType.FooterEven];
                footer?.Remove();
            }

            doc.Save(ArtifactsDir + "HeaderFooter.RemoveFooters.doc");
            //ExEnd:RemoveFooters
        }
    }
}