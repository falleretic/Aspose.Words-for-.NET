using System;
using System.Drawing;
using System.Text;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Load_Options
{
    class LoadOptionsEx : TestDataHelper
    {
        [Test]
        public static void UpdateDirtyFields()
        {
            //ExStart:UpdateDirtyFields
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.UpdateDirtyFields = true;

            Document doc = new Document(MyDir + "Dirty field.docx", loadOptions);
            doc.Save(ArtifactsDir + "LoadOptions.UpdateDirtyFields.docx");
            //ExEnd:UpdateDirtyFields
        }

        [Test]
        public static void LoadAndSaveEncryptedOdt()
        {
            //ExStart:LoadAndSaveEncryptedODT
            //ExStart:OpenEncryptedDocument
            Document doc = new Document(MyDir + "Encrypted.docx", new LoadOptions("docPassword"));
            //ExEnd:OpenEncryptedDocument
            doc.Save(ArtifactsDir + "LoadOptions.LoadAndSaveEncryptedOdt.odt", new OdtSaveOptions("newpassword"));
            //ExEnd:LoadAndSaveEncryptedODT
        }

        [Test]
        public static void ConvertShapeToOfficeMath()
        {
            //ExStart:ConvertShapeToOfficeMath
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ConvertShapeToOfficeMath = true;

            Document doc = new Document(MyDir + "Office math.docx", loadOptions);
            doc.Save(ArtifactsDir + "ConvertShapeToOfficeMath.docx", SaveFormat.Docx);
            //ExEnd:ConvertShapeToOfficeMath
        }

        [Test]
        public static void SetMsWordVersion()
        {
            //ExStart:SetMSWordVersion
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.MswVersion = MsWordVersion.Word2003;

            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            doc.Save(ArtifactsDir + "SetMsWordVersion.docx");
            //ExEnd:SetMSWordVersion
        }

        [Test]
        public static void SetTempFolder()
        {
            //ExStart:SetTempFolder  
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.TempFolder = ArtifactsDir;

            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            //ExEnd:SetTempFolder  
        }
        
        [Test]
        public static void WarningCallback()
        {
            //ExStart:WarningCallback
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.WarningCallback = new DocumentLoadingWarningCallback();
 
            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            //ExEnd:WarningCallback
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
        public static void ResourceLoadingCallback()
        {
            //ExStart:ResourceLoadingCallback
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback();
 
            // When we open an Html document, external resources such as references to CSS stylesheet files and external images
            // will be handled in a custom manner by the loading callback as the document is loaded.
            Document doc = new Document(MyDir + "Images.html", loadOptions);
            doc.Save(ArtifactsDir + "LoadOptions.ResourceLoadingCallback.pdf");
            //ExEnd:ResourceLoadingCallback
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
 
                        // CSS file will don't used in the document.
                        return ResourceLoadingAction.Skip;
                    }
                    case ResourceType.Image:
                    {
                        // Replaces all images with a substitute.
                        Image newImage = Image.FromFile(ImagesDir + "Logo.jpg");
                        ImageConverter converter = new ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(newImage, typeof(byte[]));

                        args.SetData(imageBytes);
 
                        // New images will be used instead of presented in the document.
                        return ResourceLoadingAction.UserProvided;
                    }
                    case ResourceType.Document:
                    {
                        Console.WriteLine($"External document found upon loading: {args.OriginalUri}");
 
                        // Will be used as usual.
                        return ResourceLoadingAction.Default;
                    }
                    default:
                        throw new InvalidOperationException("Unexpected ResourceType value.");
                }
            }
        }
        //ExEnd:HtmlLinkedResourceLoadingCallback

        [Test]
        public static void LoadUsingEncoding()
        {
            //ExStart:LoadUsingEncoding
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Encoding = Encoding.UTF7;

            Document doc = new Document(MyDir + "Encoded in UTF-7.txt", loadOptions);
            //ExEnd:LoadUsingEncoding
        }
    }
}