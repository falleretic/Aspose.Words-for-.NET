using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithOdtSaveOptions : SiteExamplesBase
    {
        [Test, Description("Shows how to specify unit of measure.")]
        public void MeasureUnit()
        {
            //ExStart:MeasureUnit
            Document doc = new Document(MyDir + "Document.docx");

            // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
            // and content properties in documents whereas MS Office uses inches.
            OdtSaveOptions saveOptions = new OdtSaveOptions();
            saveOptions.MeasureUnit = OdtSaveMeasureUnit.Inches;

            doc.Save(ArtifactsDir + "OdtSaveOptions.MeasureUnit.odt", saveOptions);
            //ExEnd:MeasureUnit
        }
    }
}