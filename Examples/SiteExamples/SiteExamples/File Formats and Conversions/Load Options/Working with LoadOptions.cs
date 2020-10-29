using System;
using System.Drawing;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Load_Options
{
    internal class LoadOptionsEx : SiteExamplesBase
    {
        [Test, Description("Shows how to update the dirty attribute of the fields.")]
        public void UpdateDirtyFields()
        {
            //ExStart:UpdateDirtyFields
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.UpdateDirtyFields = true;

            Document doc = new Document(MyDir + "Dirty field.docx", loadOptions);
            doc.Save(ArtifactsDir + "LoadOptions.UpdateDirtyFields.docx");
            //ExEnd:UpdateDirtyFields
        }

        [Test, Description("Shows how to load DOCX encrypted document and save as an ODT document.")]
        public void LoadAndSaveEncryptedOdt()
        {
            //ExStart:LoadAndSaveEncryptedODT
            //ExStart:OpenEncryptedDocument
            Document doc = new Document(MyDir + "Encrypted.docx", new LoadOptions("docPassword"));
            //ExEnd:OpenEncryptedDocument
            doc.Save(ArtifactsDir + "LoadOptions.LoadAndSaveEncryptedOdt.odt", new OdtSaveOptions("newpassword"));
            //ExEnd:LoadAndSaveEncryptedODT
        }

        [Test, Description("Shows how to convert shapes to OfficeMath objects during the loading document.")]
        public void ConvertShapeToOfficeMath()
        {
            //ExStart:ConvertShapeToOfficeMath
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ConvertShapeToOfficeMath = true;

            Document doc = new Document(MyDir + "Office math.docx", loadOptions);
            doc.Save(ArtifactsDir + "LoadOptions.ConvertShapeToOfficeMath.docx", SaveFormat.Docx);
            //ExEnd:ConvertShapeToOfficeMath
        }

        [Test, Description("Shows how to set MS Word version for loading document.")]
        public void SetMsWordVersion()
        {
            //ExStart:SetMSWordVersion
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.MswVersion = MsWordVersion.Word2003;

            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            doc.Save(ArtifactsDir + "LoadOptions.SetMsWordVersion.docx");
            //ExEnd:SetMSWordVersion
        }

        [Test, Description("Shows how to use a temporary folder during the loading document.")]
        public void UseTempFolder()
        {
            //ExStart:UseTempFolder  
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.TempFolder = ArtifactsDir;

            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            //ExEnd:UseTempFolder  
        }
        
        [Test, Description("Shows how to use WarningCallback to get a warnings info during the loading document.")]
        public void WarningCallback()
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
        
#if NET462
        [Test, Description("Shows how to control external resources using IResourceLoadingCallback.")]
        public void ResourceLoadingCallback()
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
#endif

        [Test, Description("Shows how to set encoding during the loading HTML/TXT documents.")]
        public void LoadUsingEncoding()
        {
            //ExStart:LoadUsingEncoding
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Encoding = Encoding.UTF7;

            Document doc = new Document(MyDir + "Encoded in UTF-7.txt", loadOptions);
            //ExEnd:LoadUsingEncoding
        }
    }
}