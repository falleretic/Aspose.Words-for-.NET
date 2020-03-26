using Aspose.Words.Fields;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractTableOfContents : TestDataHelper
    {
        public static void Run()
        {
            Document doc = new Document(DocumentDir + "TOC.doc");

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type.Equals(FieldType.FieldHyperlink))
                {
                    FieldHyperlink hyperlink = (FieldHyperlink) field;
                    if (hyperlink.SubAddress != null && hyperlink.SubAddress.StartsWith("_Toc"))
                    {
                        Paragraph tocItem = (Paragraph) field.Start.GetAncestor(NodeType.Paragraph);
                        
                        Console.WriteLine(tocItem.ToString(SaveFormat.Text).Trim());
                        Console.WriteLine("------------------");

                        Bookmark bm = doc.Range.Bookmarks[hyperlink.SubAddress];
                        // Get the location this TOC Item is pointing to
                        Paragraph pointer = (Paragraph) bm.BookmarkStart.GetAncestor(NodeType.Paragraph);
                        
                        Console.WriteLine(pointer.ToString(SaveFormat.Text));
                    }
                }
            }
        }
    }
}