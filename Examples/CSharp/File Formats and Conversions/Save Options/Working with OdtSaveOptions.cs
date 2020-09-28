using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WorkingWithOdtSaveOptions : TestDataHelper
    {
        [Test]
        public static void SetMeasureUnitForOdt()
        {
            //ExStart:SetMeasureUnitForODT
            // Load the Word document
            Document doc = new Document(DocumentDir + "Document.docx");

            // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
            // and content properties in documents whereas MS Office uses inches
            OdtSaveOptions saveOptions = new OdtSaveOptions();
            saveOptions.MeasureUnit = OdtSaveMeasureUnit.Inches;

            doc.Save(ArtifactsDir + "MeasureUnit.odt", saveOptions);
            //ExEnd:SetMeasureUnitForODT
        }
    }
}