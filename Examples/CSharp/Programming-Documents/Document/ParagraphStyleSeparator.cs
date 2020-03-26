using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ParagraphStyleSeparator : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:ParagraphStyleSeparator
            Document doc = new Document(DocumentDir + "TestFile.doc");

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