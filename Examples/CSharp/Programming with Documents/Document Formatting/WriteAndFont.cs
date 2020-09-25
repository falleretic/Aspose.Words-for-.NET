using System.Drawing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class WriteAndFont : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:WriteAndFont
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting before adding text
            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            builder.Write("Sample text.");
            
            doc.Save(ArtifactsDir + "WriteAndFont.doc");
            //ExEnd:WriteAndFont
        }
    }
}