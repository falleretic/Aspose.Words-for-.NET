using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    class WorkingWithHarfBuzz : DocsExamplesBase
    {
        [Test]
        public static void OpenTypeFeatures()
        {
            //ExStart:OpenTypeFeatures
            Document doc = new Document(MyDir + "OpenType text shaping.docx");

            // When text shaper factory is set, layout starts to use OpenType features
            // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
            doc.LayoutOptions.TextShaperFactory = Aspose.Words.Shaping.HarfBuzz.HarfBuzzTextShaperFactory.Instance;

            // Render the document to PDF format
            doc.Save(ArtifactsDir + "OpenType.Document.pdf");
            //ExEnd:OpenTypeFeatures
        }
    }
}