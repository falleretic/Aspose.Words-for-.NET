using System.Collections;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ListUseDestinationStyles : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ListUseDestinationStyles
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.DestinationList.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.SourceList.doc");

            // Set the source document to continue straight after the end of the destination document
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Keep track of the lists that are created
            Hashtable newLists = new Hashtable();

            // Iterate through all paragraphs in the document
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.IsListItem)
                {
                    int listId = para.ListFormat.List.ListId;

                    // Check if the destination document contains a list with this ID already. If it does then this may
                    // cause the two lists to run together. Create a copy of the list in the source document instead
                    if (dstDoc.Lists.GetListByListId(listId) != null)
                    {
                        Lists.List currentList;
                        // A newly copied list already exists for this ID, retrieve the stored list and use it on 
                        // the current paragraph
                        if (newLists.Contains(listId))
                        {
                            currentList = (Lists.List) newLists[listId];
                        }
                        else
                        {
                            // Add a copy of this list to the document and store it for later reference
                            currentList = srcDoc.Lists.AddCopy(para.ListFormat.List);
                            newLists.Add(listId, currentList);
                        }

                        // Set the list of this paragraph  to the copied list
                        para.ListFormat.List = currentList;
                    }
                }
            }

            // Append the source document to end of the destination document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            dstDoc.Save(ArtifactsDir + "ListUseDestinationStyles.docx");
            //ExEnd:ListUseDestinationStyles
        }
    }
}