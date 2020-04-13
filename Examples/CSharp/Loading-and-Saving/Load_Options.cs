using Aspose.Words.Saving;
using Aspose.Words.Settings;
using System;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class Load_Options : TestDataHelper
    {
        public static void Run()
        {
            LoadOptionsUpdateDirtyFields();
            LoadAndSaveEncryptedOdt();
            VerifyOdtDocument();
            ConvertShapeToOfficeMath();
            SetMsWordVersion();
        }

        public static void LoadOptionsUpdateDirtyFields()
        {
            //ExStart:LoadOptionsUpdateDirtyFields
            LoadOptions lo = new LoadOptions();
            // Update the fields with the dirty attribute
            lo.UpdateDirtyFields = true;

            Document doc = new Document(LoadingSavingDir + "input.docx", lo);
            doc.Save(ArtifactsDir + "LoadOptionsUpdateDirtyFields.docx");
            //ExEnd:LoadOptionsUpdateDirtyFields
        }

        public static void LoadAndSaveEncryptedOdt()
        {
            //ExStart:LoadAndSaveEncryptedODT
            Document doc = new Document(LoadingSavingDir + "encrypted.odt", new LoadOptions("password"));
            doc.Save(ArtifactsDir + "LoadAndSaveEncryptedOdt.odt", new OdtSaveOptions("newpassword"));
            //ExEnd:LoadAndSaveEncryptedODT
        }

        public static void VerifyOdtDocument()
        {
            //ExStart:VerifyODTdocument
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(LoadingSavingDir + "encrypted.odt");
            Console.WriteLine(info.IsEncrypted);
            //ExEnd:VerifyODTdocument
        }

        public static void ConvertShapeToOfficeMath()
        {
            //ExStart:ConvertShapeToOfficeMath
            LoadOptions lo = new LoadOptions();
            lo.ConvertShapeToOfficeMath = true;

            // Specify load option to use previous default behaviour i.e. convert math shapes to office math ojects on loading stage.
            Document doc = new Document(LoadingSavingDir + "OfficeMath.docx", lo);
            doc.Save(ArtifactsDir + "ConvertShapeToOfficeMath.docx", SaveFormat.Docx);
            //ExEnd:ConvertShapeToOfficeMath
        }

        public static void SetMsWordVersion()
        {
            //ExStart:SetMSWordVersion
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.MswVersion = MsWordVersion.Word2003;
            Document doc = new Document(LoadingSavingDir + "document.doc", loadOptions);

            doc.Save(ArtifactsDir + "SetMsWordVersion.docx");
            //ExEnd:SetMSWordVersion
        }
    }
}