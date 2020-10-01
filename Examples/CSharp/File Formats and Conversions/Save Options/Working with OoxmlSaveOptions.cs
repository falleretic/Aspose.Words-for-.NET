using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class WorkingWithOoxmlSaveOptions : TestDataHelper
    {
        [Test]
        public static void EncryptDocxWithPassword()
        {
            //ExStart:EncryptDocxWithPassword
            Document doc = new Document(MyDir + "Document.docx");
            
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "OoxmlSaveOptions.EncryptDocxWithPassword.docx", saveOptions);
            //ExEnd:EncryptDocxWithPassword
        }

        [Test]
        public static void OoxmlComplianceIso29500_2008_Strict()
        {
            //ExStart:OoxmlComplianceIso29500_2008_Strict
            Document doc = new Document(MyDir + "Document.docx");

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            saveOptions.SaveFormat = SaveFormat.Docx;

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.OoxmlComplianceIso29500_2008_Strict.docx", saveOptions);
            //ExEnd:OoxmlComplianceIso29500_2008_Strict
        }

        [Test]
        public static void UpdateLastSavedTimeProperty()
        {
            //ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(MyDir + "Document.docx");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateLastSavedTimeProperty = true;

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.UpdateLastSavedTimeProperty.docx", saveOptions);
            //ExEnd:UpdateLastSavedTimeProperty
        }

        [Test]
        public static void KeepLegacyControlChars()
        {
            //ExStart:KeepLegacyControlChars
            Document doc = new Document(MyDir + "Legacy control character.doc");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.FlatOpc);
            saveOptions.KeepLegacyControlChars = true;

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.KeepLegacyControlChars.docx", saveOptions);
            //ExEnd:KeepLegacyControlChars
        }

        [Test]
        public static void SetCompressionLevel()
        {
            // ExStart:SetCompressionLevel
            Document doc = new Document(MyDir + "Document.docx");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            saveOptions.CompressionLevel = CompressionLevel.SuperFast;

            doc.Save(ArtifactsDir + "OoxmlSaveOptionsEx.SetCompressionLevel.docx", saveOptions);
            // ExEnd:SetCompressionLevel
        }
    }
}