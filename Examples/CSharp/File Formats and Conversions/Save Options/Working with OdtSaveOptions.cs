using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class WorkingWithOdtSaveOptions : TestDataHelper
    {
        [Test]
        public static void MeasureUnit()
        {
            //ExStart:SetMeasureUnitForODT
            Document doc = new Document(MyDir + "Document.docx");

            // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
            // and content properties in documents whereas MS Office uses inches.
            OdtSaveOptions saveOptions = new OdtSaveOptions();
            saveOptions.MeasureUnit = OdtSaveMeasureUnit.Inches;

            doc.Save(ArtifactsDir + "OdtSaveOptions.MeasureUnit.odt", saveOptions);
            //ExEnd:SetMeasureUnitForODT
        }
    }
}