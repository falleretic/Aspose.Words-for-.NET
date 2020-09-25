using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.OpenType
{
    class WorkingWithHarfBuzz : TestDataHelper
    {
        [Test]
        public static void OpenTypeFeatures()
        {
            //ExStart:OpenTypeFeatures
            Document doc = new Document(DocumentDir + "OpenType text shaping.docx");

            // When text shaper factory is set, layout starts to use OpenType features
            // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
            doc.LayoutOptions.TextShaperFactory = Shaping.HarfBuzz.HarfBuzzTextShaperFactory.Instance;

            // Render the document to PDF format
            doc.Save(ArtifactsDir + "OpenType.Document.pdf");
            //ExEnd:OpenTypeFeatures
        }
    }
}