using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithOoxml : TestDataHelper
    {
        [Test]
        public static void EncryptDocxWithPassword()
        {
            //ExStart:EncryptDocxWithPassword
            Document doc = new Document(LoadingSavingDir + "Document.doc");
            
            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "EncryptDocxWithPassword.docx", ooxmlSaveOptions);
            //ExEnd:EncryptDocxWithPassword
        }

        [Test]
        public static void SetOoxmlCompliance()
        {
            //ExStart:SetOOXMLCompliance
            Document doc = new Document(LoadingSavingDir + "Document.doc");

            // Set Word2016 version for document
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            // Set the Strict compliance level. 
            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            ooxmlSaveOptions.SaveFormat = SaveFormat.Docx;

            doc.Save(ArtifactsDir + "SetOoxmlCompliance.docx", ooxmlSaveOptions);
            //ExEnd:SetOOXMLCompliance
        }

        [Test]
        public static void UpdateLastSavedTimeProperty()
        {
            //ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(LoadingSavingDir + "Document.doc");

            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.UpdateLastSavedTimeProperty = true;

            doc.Save(ArtifactsDir + "UpdateLastSavedTimeProperty.docx", ooxmlSaveOptions);
            //ExEnd:UpdateLastSavedTimeProperty
        }

        [Test]
        public static void KeepLegacyControlChars()
        {
            //ExStart:KeepLegacyControlChars
            Document doc = new Document(LoadingSavingDir + "Document.doc");

            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.FlatOpc);
            so.KeepLegacyControlChars = true;

            doc.Save(ArtifactsDir + "KeepLegacyControlChars.docx", so);
            //ExEnd:KeepLegacyControlChars
        }

        [Test]
        public static void SetCompressionLevel()
        {
            // ExStart:SetCompressionLevel
            Document doc = new Document(LoadingSavingDir + "Document.doc");

            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.Docx);
            so.CompressionLevel = CompressionLevel.SuperFast;

            // Save the document to disk.
            doc.Save(ArtifactsDir + "SetCompressionLevel.docx", so);
            // ExEnd:SetCompressionLevel
        }
    }
}