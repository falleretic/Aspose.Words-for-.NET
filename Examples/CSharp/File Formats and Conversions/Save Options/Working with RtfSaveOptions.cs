using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class WorkingWithRtfSaveOptions : TestDataHelper
    {
        [Test]
        public static void SavingImagesAsWmf()
        {
            //ExStart:SavingImagesAsWmf
            Document doc = new Document(MyDir + "Document.docx");

            RtfSaveOptions saveOptions = new RtfSaveOptions();
            saveOptions.SaveImagesAsWmf = true;

            doc.Save(ArtifactsDir + "RtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
            //ExEnd:SavingImagesAsWmf
        }
    }
}