using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithOoxmlSaveOptions : SiteExamplesBase
    {
        [Test, Description("Shows how to encrypt document with password.")]
        public void EncryptDocxWithPassword()
        {
            //ExStart:EncryptDocxWithPassword
            Document doc = new Document(MyDir + "Document.docx");
            
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.EncryptDocxWithPassword.docx", saveOptions);
            //ExEnd:EncryptDocxWithPassword
        }

        [Test, Description("Shows how to specify OOXML version.")]
        public void OoxmlComplianceIso29500_2008_Strict()
        {
            //ExStart:OoxmlComplianceIso29500_2008_Strict
            Document doc = new Document(MyDir + "Document.docx");

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            saveOptions.SaveFormat = SaveFormat.Docx;

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.OoxmlComplianceIso29500_2008_Strict.docx", saveOptions);
            //ExEnd:OoxmlComplianceIso29500_2008_Strict
        }

        [Test, Description("Shows how to update last saved time before saving.")]
        public void UpdateLastSavedTimeProperty()
        {
            //ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(MyDir + "Document.docx");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateLastSavedTimeProperty = true;

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.UpdateLastSavedTimeProperty.docx", saveOptions);
            //ExEnd:UpdateLastSavedTimeProperty
        }

        [Test, Description("Shows how to keep original representation of legacy control chars.")]
        public void KeepLegacyControlChars()
        {
            //ExStart:KeepLegacyControlChars
            Document doc = new Document(MyDir + "Legacy control character.doc");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.FlatOpc);
            saveOptions.KeepLegacyControlChars = true;

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.KeepLegacyControlChars.docx", saveOptions);
            //ExEnd:KeepLegacyControlChars
        }

        [Test, Description("Shows how to specify compression level.")]
        public void SetCompressionLevel()
        {
            // ExStart:SetCompressionLevel
            Document doc = new Document(MyDir + "Document.docx");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            saveOptions.CompressionLevel = CompressionLevel.SuperFast;

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.SetCompressionLevel.docx", saveOptions);
            // ExEnd:SetCompressionLevel
        }
    }
}