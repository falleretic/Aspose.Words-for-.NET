using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithOoxmlSaveOptionsEx : TestDataHelper
    {
        [Test]
        public static void EncryptDocxWithPassword()
        {
            //ExStart:EncryptDocxWithPassword
            Document doc = new Document(MyDir + "Document.docx");
            
            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.Password = "password";
            
            doc.Save(ArtifactsDir + "OoxmlSaveOptionsEx.EncryptDocxWithPassword.docx", ooxmlSaveOptions);
            //ExEnd:EncryptDocxWithPassword
        }

        [Test]
        public static void SetOoxmlCompliance()
        {
            //ExStart:SetOOXMLCompliance
            Document doc = new Document(LoadingSavingDir + "Document.docx");

            // Set Word2016 version for document
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            // Set the Strict compliance level. 
            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            ooxmlSaveOptions.SaveFormat = SaveFormat.Docx;

            doc.Save(ArtifactsDir + "OoxmlSaveOptionsEx.SetOoxmlCompliance.docx", ooxmlSaveOptions);
            //ExEnd:SetOOXMLCompliance
        }

        [Test]
        public static void UpdateLastSavedTimeProperty()
        {
            //ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(LoadingSavingDir + "Document.docx");

            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.UpdateLastSavedTimeProperty = true;

            doc.Save(ArtifactsDir + "OoxmlSaveOptionsEx.UpdateLastSavedTimeProperty.docx", ooxmlSaveOptions);
            //ExEnd:UpdateLastSavedTimeProperty
        }

        [Test]
        public static void KeepLegacyControlChars()
        {
            //ExStart:KeepLegacyControlChars
            Document doc = new Document(LoadingSavingDir + "Legacy control character.doc");

            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions(SaveFormat.FlatOpc);
            ooxmlSaveOptions.KeepLegacyControlChars = true;

            doc.Save(ArtifactsDir + "OoxmlSaveOptionsEx.KeepLegacyControlChars.docx", ooxmlSaveOptions);
            //ExEnd:KeepLegacyControlChars
        }

        [Test]
        public static void SetCompressionLevel()
        {
            // ExStart:SetCompressionLevel
            Document doc = new Document(LoadingSavingDir + "Document.docx");

            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            ooxmlSaveOptions.CompressionLevel = CompressionLevel.SuperFast;

            // Save the document to disk.
            doc.Save(ArtifactsDir + "OoxmlSaveOptionsEx.SetCompressionLevel.docx", ooxmlSaveOptions);
            // ExEnd:SetCompressionLevel
        }
    }
}