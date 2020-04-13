using System.IO;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SpecifySaveOption : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:SpecifySaveOption
            Document doc = new Document(LoadingSavingDir + "TestFile RenderShape.docx");

            // This is the directory we want the exported images to be saved to
            string imagesDir = Path.Combine(ArtifactsDir, "Images");

            // The folder specified needs to exist and should be empty
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportTextInputFormFieldAsText = true;
            options.ImagesFolder = imagesDir;

            doc.Save(ArtifactsDir + "SpecifySaveOption.html", options);
            //ExEnd:SpecifySaveOption
        }
    }
}