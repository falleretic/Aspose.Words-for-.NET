using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    public class WorkingWithSaveOptions : TestDataHelper
    {
        public static void Run()
        {
            UpdateLastSavedTimeProperty();
            SetMeasureUnitForOdt();
        }

        public static void UpdateLastSavedTimeProperty()
        {
            //ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(DocumentDir + "Document.doc");

            OoxmlSaveOptions options = new OoxmlSaveOptions();
            options.UpdateLastSavedTimeProperty = true;

            doc.Save(ArtifactsDir + "UpdateLastSavedTimeProperty.docx", options);
            //ExEnd:UpdateLastSavedTimeProperty
        }

        public static void SetMeasureUnitForOdt()
        {
            //ExStart:SetMeasureUnitForODT
            // Load the Word document
            Document doc = new Document(DocumentDir + "Document.doc");

            // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
            // and content properties in documents whereas MS Office uses inches
            OdtSaveOptions saveOptions = new OdtSaveOptions();
            saveOptions.MeasureUnit = OdtSaveMeasureUnit.Inches;

            doc.Save(ArtifactsDir + "MeasureUnit.odt", saveOptions);
            //ExEnd:SetMeasureUnitForODT
        }
    }
}