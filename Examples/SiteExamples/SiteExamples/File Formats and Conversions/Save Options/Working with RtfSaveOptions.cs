using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Save_Options
{
    class WorkingWithRtfSaveOptions : SiteExamplesBase
    {
        [Test, Description("Shows how to save images as WMF.")]
        public static void SavingImagesAsWmf()
        {
            //ExStart:SavingImagesAsWmf
            Document doc = new Document(MyDir + "Document.docx");

            RtfSaveOptions saveOptions = new RtfSaveOptions();
            saveOptions.SaveImagesAsWmf = true;

            doc.Save(ArtifactsDir + "WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
            //ExEnd:SavingImagesAsWmf
        }
    }
}