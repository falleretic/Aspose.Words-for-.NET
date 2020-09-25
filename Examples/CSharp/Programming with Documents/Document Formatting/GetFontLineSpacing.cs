using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class GetFontLineSpacing
    {
        [Test]
        public static void Run()
        {
            //ExStart:GetFontLineSpacing
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Font.Name = "Calibri";
            builder.Writeln("qText");

            // Obtain line spacing
            Font font = builder.Document.FirstSection.Body.FirstParagraph.Runs[0].Font;
            Console.WriteLine($"lineSpacing = {font.LineSpacing}");
            //ExEnd:GetFontLineSpacing
        }
    }
}