using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    internal class RemoveWatermark : TestDataHelper
    {
        [Test]
        //ExStart:RemoveWatermark
        public static void Run()
        {
            Document doc = new Document(ImagesDir + "RemoveWatermark.docx");
            
            RemoveWatermarkText(doc);
            
            doc.Save(ArtifactsDir + "RemoveWatermark.docx");
        }

        private static void RemoveWatermarkText(Document doc)
        {
            foreach (HeaderFooter hf in doc.GetChildNodes(NodeType.HeaderFooter, true))
            {
                foreach (Shape shape in hf.GetChildNodes(NodeType.Shape, true))
                {
                    if (shape.Name.Contains("WaterMark"))
                    {
                        shape.Remove();
                    }
                }
            }
        }
    }
    //ExEnd:RemoveWatermark
}