using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    class ExtractImagesToFiles : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ExtractImagesToFiles
            Document doc = new Document(ImagesDir + "Image.SampleImages.doc");

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