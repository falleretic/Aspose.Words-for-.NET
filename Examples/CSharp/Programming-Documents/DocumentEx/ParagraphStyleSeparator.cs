using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class ParagraphStyleSeparator : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ParagraphStyleSeparator
            Document doc = new Document(DocumentDir + "Document.docx");

            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (paragraph.BreakIsStyleSeparator)
                {
                    Console.WriteLine("Separator Found!");
                }
            }
            //ExEnd:ParagraphStyleSeparator
        }
    }
}