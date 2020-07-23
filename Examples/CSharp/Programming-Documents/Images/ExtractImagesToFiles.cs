using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Images
{
    class ExtractImagesToFiles : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExtractImagesToFiles
            Document doc = new Document(ImagesDir + "Images.docx");

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = string.Format(
                        "Image.ExportImages.{0}_out{1}", imageIndex,
                        FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType));

                    shape.ImageData.Save(ArtifactsDir + imageFileName);
                    imageIndex++;
                }
            }
            //ExEnd:ExtractImagesToFiles
        }
    }
}