using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Load_Options
{
    class WorkingWithRtfLoadOptions : TestDataHelper
    {
        [Test]
        public static void RecognizeUtf8Text()
        {
            //ExStart:RecognizeUtf8Text
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            loadOptions.RecognizeUtf8Text = true;

            Document doc = new Document(MyDir + "UTF-8 characters.rtf", loadOptions);
            doc.Save(ArtifactsDir + "RtfLoadOptions.RecognizeUtf8Text.rtf");
            //ExEnd:RecognizeUtf8Text
        }
    }
}