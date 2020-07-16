using System;
using System.Collections;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class RemoveTOCFromDocument : TestDataHelper
    {
        [Test]
        //ExStart:RemoveTOCFromDocument
        public static void Run()
        {
            Document doc = new Document(StyleDir + "Document.TableOfContents.doc");

            // Remove the first table of contents from the document
            RemoveTableOfContents(doc, 0);

            doc.Save(ArtifactsDir + "Document.TableOfContentsRemoveToc.doc");
        }

        /// <summary>
        /// Removes the specified table of contents field from the document.
        /// </summary>
        /// <param name="doc">The document to remove the field from.</param>
        /// <param name="index">The zero-based index of the TOC to remove.</param>
        public static void RemoveTableOfContents(Document doc, int index)
        {
            // Store the FieldStart nodes of TOC fields in the document for quick access
            ArrayList fieldStarts = new ArrayList();
            // This is a list to store the nodes found inside the specified TOC. They will be removed
            // At the end of this method
            ArrayList nodeList = new ArrayList();

            foreach (FieldStart start in doc.GetChildNodes(NodeType.FieldStart, true))
            {
                if (start.FieldType == FieldType.FieldTOC)
                {
                    // Add all FieldStarts which are of type FieldTOC
                    fieldStarts.Add(start);
                }
            }

            // Ensure the TOC specified by the passed index exists
            if (index > fieldStarts.Count - 1)
                throw new ArgumentOutOfRangeException("TOC index is out of range");

            bool isRemoving = true;
            // Get the FieldStart of the specified TOC
            Node currentNode = (Node) fieldStarts[index];

            while (isRemoving)
            {
                // It is safer to store these nodes and delete them all at once later
                nodeList.Add(currentNode);
                currentNode = currentNode.NextPreOrder(doc);

                // Once we encounter a FieldEnd node of type FieldTOC then we know we are at the end
                // Of the current TOC and we can stop here
                if (currentNode.NodeType == NodeType.FieldEnd)
                {
                    FieldEnd fieldEnd = (FieldEnd) currentNode;
                    if (fieldEnd.FieldType == FieldType.FieldTOC)
                        isRemoving = false;
                }
            }

            // Remove all nodes found in the specified TOC
            foreach (Node node in nodeList)
            {
                node.Remove();
            }
        }
        //ExEnd:RemoveTOCFromDocument
    }
}