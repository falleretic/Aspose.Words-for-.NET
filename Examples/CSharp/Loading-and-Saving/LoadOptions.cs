using Aspose.Words.Saving;
using Aspose.Words.Settings;
using System;
using System.Drawing;
using System.Text;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class LoadOptions : TestDataHelper
    {
        [Test]
        public static void LoadOptionsUpdateDirtyFields()
        {
            //ExStart:LoadOptionsUpdateDirtyFields
            Words.LoadOptions lo = new Words.LoadOptions();
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
            Document doc = new Document(QuickStartDir + "encrypted.odt", new Words.LoadOptions("password"));
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
            Words.LoadOptions lo = new Words.LoadOptions();
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
            Words.LoadOptions loadOptions = new Words.LoadOptions();
            loadOptions.MswVersion = MsWordVersion.Word2003;
            Document doc = new Document(LoadingSavingDir + "document.doc", loadOptions);

            doc.Save(ArtifactsDir + "SetMsWordVersion.docx");
            //ExEnd:SetMSWordVersion
        }

        [Test]
        public static void SetTempFolder()
        {
            // ExStart:SetTempFolder  
            Words.LoadOptions lo = new Words.LoadOptions();
            lo.TempFolder = ArtifactsDir;

            Document doc = new Document(LoadingSavingDir + "document.docx", lo);
            // ExEnd:SetTempFolder  
        }
        
        [Test]
        public static void LoadOptionsWarningCallback()
        {
            //ExStart:LoadOptionsWarningCallback
            // Create a new LoadOptions object and set its WarningCallback property. 
            Words.LoadOptions loadOptions = new Words.LoadOptions { WarningCallback = new DocumentLoadingWarningCallback() };
 
            Document doc = new Document(LoadingSavingDir + "document.docx", loadOptions);
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
        
        [Test]
        public static void LoadOptionsResourceLoadingCallback()
        {
            //ExStart:LoadOptionsResourceLoadingCallback
            // Create a new LoadOptions object and set its ResourceLoadingCallback attribute as an instance of our IResourceLoadingCallback implementation
            Words.LoadOptions loadOptions = new Words.LoadOptions { ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback() };
 
            // When we open an Html document, external resources such as references to CSS stylesheet files and external images
            // will be handled in a custom manner by the loading callback as the document is loaded
            Document doc = new Document(LoadingSavingDir + "Images.html", loadOptions);
            doc.Save(ArtifactsDir + "Document.LoadOptionsCallback.pdf");
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
                        Image newImage = Image.FromFile(QuickStartDir + "Logo.jpg");
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

        [Test]
        public static void LoadOptionsEncoding()
        {
            //ExStart:LoadOptionsEncoding
            // Set the Encoding attribute in a LoadOptions object to override the automatically chosen encoding with the one we know to be correct
            Words.LoadOptions loadOptions = new Words.LoadOptions { Encoding = Encoding.UTF7 };
            Document doc = new Document(LoadingSavingDir + "Encoded in UTF-7.txt", loadOptions);
            //ExEnd:LoadOptionsEncoding
        }
    }
}