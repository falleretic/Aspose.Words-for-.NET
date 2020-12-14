using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    internal class WorkingWithHarfBuzz : DocsExamplesBase
    {
        [Test]
        public static void OpenTypeFeatures()
        {
            //ExStart:OpenTypeFeatures
            Document doc = new Document(MyDir + "OpenType text shaping.docx");

            // When we set the text shaper factory, the layout starts to use OpenType features.
            // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory.
            doc.LayoutOptions.TextShaperFactory = Aspose.Words.Shaping.HarfBuzz.HarfBuzzTextShaperFactory.Instance;

            doc.Save(ArtifactsDir + "WorkingWithHarfBuzz.OpenTypeFeatures.pdf");
            //ExEnd:OpenTypeFeatures
        }
    }
}