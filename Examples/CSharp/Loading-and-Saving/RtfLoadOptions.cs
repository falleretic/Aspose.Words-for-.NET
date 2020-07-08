using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithRTF : TestDataHelper
    {
        [Test]
        public static void RecognizeUtf8Text()
        {
            //ExStart:RecognizeUtf8Text
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            loadOptions.RecognizeUtf8Text = true;

            Document doc = new Document(LoadingSavingDir + "Utf8Text.rtf", loadOptions);
            doc.Save(ArtifactsDir + "RecognizeUtf8Text.rtf");
            //ExEnd:RecognizeUtf8Text
        }
    }
}