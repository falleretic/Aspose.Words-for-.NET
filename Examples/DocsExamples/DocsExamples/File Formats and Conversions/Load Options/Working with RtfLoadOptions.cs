using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Load_Options
{
    internal class WorkingWithRtfLoadOptions : DocsExamplesBase
    {
        [Test]
        public void RecognizeUtf8Text()
        {
            //ExStart:RecognizeUtf8Text
            RtfLoadOptions rtfLoadOptions = new RtfLoadOptions { RecognizeUtf8Text = true };
            
            Document doc = new Document(MyDir + "UTF-8 characters.rtf", rtfLoadOptions);

            doc.Save(ArtifactsDir + "WorkingWithRtfLoadOptions.RecognizeUtf8Text.rtf");
            //ExEnd:RecognizeUtf8Text
        }
    }
}