using System;
using Aspose.Words.Layout;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class PageNumbersOfNodes : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            Document doc = new Document(DocumentDir + "Document.docx");

            // Create and attach collector before the document before page layout is built
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // This will build layout model and collect necessary information
            doc.UpdatePageLayout();

            // Print the details of each document node including the page numbers
            foreach (Node node in doc.FirstSection.Body.GetChildNodes(NodeType.Any, true))
            {
                Console.WriteLine(" --------- ");
                Console.WriteLine("NodeType:   " + Node.NodeTypeToString(node.NodeType));
                Console.WriteLine("Text:       \"" + node.ToString(SaveFormat.Text).Trim() + "\"");
                Console.WriteLine("Page Start: " + layoutCollector.GetStartPageIndex(node));
                Console.WriteLine("Page End:   " + layoutCollector.GetEndPageIndex(node));
                Console.WriteLine(" --------- ");
                Console.WriteLine();
            }

            // Detatch the collector from the document
            layoutCollector.Document = null;
        }
    }
}