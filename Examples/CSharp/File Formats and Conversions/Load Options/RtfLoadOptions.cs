using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithRTF : TestDataHelper
    {
        [Test]
        public static void RecognizeUtf8Text()
        {
            //ExStart:RecognizeUtf8Text
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            loadOptions.RecognizeUtf8Text = true;

            Document doc = new Document(LoadingSavingDir + "UTF-8 characters.rtf", loadOptions);
            doc.Save(ArtifactsDir + "RecognizeUtf8Text.rtf");
            //ExEnd:RecognizeUtf8Text
        }
    }
}