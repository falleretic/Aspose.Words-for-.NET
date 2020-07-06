using Aspose.Words.Saving;
using Aspose.Words.Settings;
using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class Load_Options : TestDataHelper
    {
        [Test]
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

        [Test]
        public static void LoadAndSaveEncryptedOdt()
        {
            //ExStart:LoadAndSaveEncryptedODT
            Document doc = new Document(QuickStartDir + "encrypted.odt", new LoadOptions("password"));
            doc.Save(ArtifactsDir + "LoadAndSaveEncryptedOdt.odt", new OdtSaveOptions("newpassword"));
            //ExEnd:LoadAndSaveEncryptedODT
        }

        [Test]
        public static void VerifyOdtDocument()
        {
            //ExStart:VerifyODTdocument
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(QuickStartDir + "encrypted.odt");
            Console.WriteLine(info.IsEncrypted);
            //ExEnd:VerifyODTdocument
        }

        [Test]
        public static void ConvertShapeToOfficeMath()
        {
            //ExStart:ConvertShapeToOfficeMath
            LoadOptions lo = new LoadOptions();
            lo.ConvertShapeToOfficeMath = true;

            // Specify load option to use previous default behaviour i.e. convert math shapes to office math ojects on loading stage.
            Document doc = new Document(QuickStartDir + "OfficeMath.docx", lo);
            doc.Save(ArtifactsDir + "ConvertShapeToOfficeMath.docx", SaveFormat.Docx);
            //ExEnd:ConvertShapeToOfficeMath
        }

        [Test]
        public static void SetMsWordVersion()
        {
            //ExStart:SetMSWordVersion
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.MswVersion = MsWordVersion.Word2003;
            Document doc = new Document(LoadingSavingDir + "document.doc", loadOptions);

            doc.Save(ArtifactsDir + "SetMsWordVersion.docx");
            //ExEnd:SetMSWordVersion
        }

        public static void SetTempFolder(string dataDir)
        {
            // ExStart:SetTempFolder  
            LoadOptions lo = new LoadOptions();
            lo.TempFolder = @"C:\TempFolder\";

            Document doc = new Document(dataDir + "document.docx", lo);
            // ExEnd:SetTempFolder  
        }
        
        public static void LoadOptionsWarningCallback(string dataDir)
        {
            //ExStart:LoadOptionsWarningCallback
            // Create a new LoadOptions object and set its WarningCallback property. 
            LoadOptions loadOptions = new LoadOptions { WarningCallback = new DocumentLoadingWarningCallback() };
 
            Document doc = new Document(dataDir + "document.docx", loadOptions);
            //ExEnd:LoadOptionsWarningCallback
        }

        //ExStart:DocumentLoadingWarningCallback
        public class DocumentLoadingWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                // Prints warnings and their details as they arise during document loading.
                Console.WriteLine($"WARNING: {info.WarningType}, source: {info.Source}");
                Console.WriteLine($"\tDescription: {info.Description}");
            }
        }
        //ExEnd:DocumentLoadingWarningCallback
        
        public static void LoadOptionsResourceLoadingCallback(string dataDir)
        {
            //ExStart:LoadOptionsResourceLoadingCallback
            // Create a new LoadOptions object and set its ResourceLoadingCallback attribute as an instance of our IResourceLoadingCallback implementation
            LoadOptions loadOptions = new LoadOptions { ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback() };
 
            // When we open an Html document, external resources such as references to CSS stylesheet files and external images
            // will be handled in a custom manner by the loading callback as the document is loaded
            Document doc = new Document(dataDir + "Images.html", loadOptions);
            doc.Save(dataDir + "Document.LoadOptionsCallback_out.pdf");
            //ExEnd:LoadOptionsResourceLoadingCallback
        }

        //ExStart:HtmlLinkedResourceLoadingCallback
        private class HtmlLinkedResourceLoadingCallback : IResourceLoadingCallback
        {
            public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
            {
                switch (args.ResourceType)
                {
                    case ResourceType.CssStyleSheet:
                    {
                        Console.WriteLine($"External CSS Stylesheet found upon loading: {args.OriginalUri}");
 
                        // CSS file will don't used in the document
                        return ResourceLoadingAction.Skip;
                    }
                    case ResourceType.Image:
                    {
                        // Replaces all images with a substitute
                        const string newImageFilename = "Logo.jpg";
                        Console.WriteLine($"\tImage will be substituted with: {newImageFilename}");
                        Image newImage = Image.FromFile(RunExamples.GetDataDir_QuickStart() + newImageFilename);
                        ImageConverter converter = new ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(newImage, typeof(byte[]));
                        args.SetData(imageBytes);
 
                        // New images will be used instead of presented in the document
                        return ResourceLoadingAction.UserProvided;
                    }
                    case ResourceType.Document:
                    {
                        Console.WriteLine($"External document found upon loading: {args.OriginalUri}");
 
                        // Will be used as usual
                        return ResourceLoadingAction.Default;
                    }
                    default:
                        throw new InvalidOperationException("Unexpected ResourceType value.");
                }
            }
        }
        //ExEnd:HtmlLinkedResourceLoadingCallback

        public static void LoadOptionsEncoding(string dataDir)
        {
            //ExStart:LoadOptionsEncoding
            // Set the Encoding attribute in a LoadOptions object to override the automatically chosen encoding with the one we know to be correct
            LoadOptions loadOptions = new LoadOptions { Encoding = Encoding.UTF7 };
            Document doc = new Document(dataDir + "Encoded in UTF-7.txt", loadOptions);
            //ExEnd:LoadOptionsEncoding
        }
    }
}