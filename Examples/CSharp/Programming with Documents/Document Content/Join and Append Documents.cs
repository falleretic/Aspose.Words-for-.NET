using System.Collections;
using System.Text;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Document_Content
{
    internal class JoinAndAppendDocuments : TestDataHelper
    {
        [Test]
        public static void SimpleAppendDocument()
        {
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Append the source document to the destination document using no extra options
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "SimpleAppendDocument.docx");
        }

        [Test]
        public static void AppendDocument()
        {
            //ExStart:AppendDocumentManually
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");
            
            // Loop through all sections in the source document.
            // Section nodes are immediate children of the Document node so we can just enumerate the Document.
            foreach (Section srcSection in srcDoc)
            {
                // Because we are copying a section from one document to another, 
                // it is required to import the Section node into the destination document.
                // This adjusts any document-specific references to styles, lists, etc.
                //
                // Importing a node creates a copy of the original node, but the copy
                // ss ready to be inserted into the destination document.
                Node dstSection = dstDoc.ImportNode(srcSection, true, ImportFormatMode.KeepSourceFormatting);

                // Now the new section node can be appended to the destination document.
                dstDoc.AppendChild(dstSection);
            }

            dstDoc.Save(ArtifactsDir + "AppendDocumentManually.docx");
            //ExEnd:AppendDocumentManually
        }

        [Test]
        public static void BaseDocument()
        {
            //ExStart:BaseDocument
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document();
            
            // The destination document is not actually empty which often causes a blank page to appear before the appended document.
            // This is due to the base document having an empty section and the new document being started on the next page.
            // Remove all content from the destination document before appending.
            dstDoc.RemoveAllChildren();
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "BaseDocument.docx");
            //ExEnd:BaseDocument
        }

        [Test]
        public static void AppendWithImportFormatOptions()
        {
            //ExStart:AppendWithImportFormatOptions
            Document srcDoc = new Document(MyDir + "Document source with list.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");

            ImportFormatOptions options = new ImportFormatOptions();
            // Specify that if numbering clashes in source and destination documents,
            // then a numbering from the source document will be used.
            options.KeepSourceNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:AppendWithImportFormatOptions
        }

        [Test]
        public static void ConvertNumPageFields()
        {
            //ExStart:ConvertNumPageFields
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Restart the page numbering on the start of the source document.
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;
            srcDoc.FirstSection.PageSetup.PageStartingNumber = 1;

            // Append the source document to the end of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // After joining the documents the NUMPAGE fields will now display the total number of pages which
            // is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields,
            ConvertNumPageFieldsToPageRef(dstDoc);

            // This needs to be called in order to update the new fields with page numbers.
            dstDoc.UpdatePageLayout();

            dstDoc.Save(ArtifactsDir + "ConvertNumPageFields.docx");
            //ExEnd:ConvertNumPageFields
        }

        //ExStart:ConvertNumPageFieldsToPageRef
        public static void ConvertNumPageFieldsToPageRef(Document doc)
        {
            // This is the prefix for each bookmark which signals where page numbering restarts.
            // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
            const string bookmarkPrefix = "_SubDocumentEnd";
            const string numPagesFieldName = "NUMPAGES";
            const string pageRefFieldName = "PAGEREF";

            DocumentBuilder builder = new DocumentBuilder(doc);
            // Defines the number of page restarts that have been encountered and therefore the number of "sub" documents
            // found within this document.
            int subDocumentCount = 0;

            foreach (Section section in doc.Sections)
            {
                // This section has it's page numbering restarted so we will treat this as the start of a sub document.
                // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
                if (section.PageSetup.RestartPageNumbering)
                {
                    // Don't do anything if this is the first section in the document. This part of the code will insert the bookmark marking
                    // the end of the previous sub-document so, therefore, it does not apply to the first section in the document.
                    if (!section.Equals(doc.FirstSection))
                    {
                        // Get the previous section and the last node within the body of that section
                        Section prevSection = (Section) section.PreviousSibling;
                        Node lastNode = prevSection.Body.LastChild;

                        // Use the DocumentBuilder to move to this node and insert the bookmark there
                        // This bookmark represents the end of the sub document
                        builder.MoveTo(lastNode);
                        builder.StartBookmark(bookmarkPrefix + subDocumentCount);
                        builder.EndBookmark(bookmarkPrefix + subDocumentCount);

                        // Increase the subdocument count to insert the correct bookmarks
                        subDocumentCount++;
                    }
                }

                // The last section simply needs the ending bookmark to signal that it is the end of the current sub document
                if (section.Equals(doc.LastSection))
                {
                    // Insert the bookmark at the end of the body of the last section
                    // Don't increase the count this time as we are just marking the end of the document
                    Node lastNode = doc.LastSection.Body.LastChild;
                    builder.MoveTo(lastNode);
                    builder.StartBookmark(bookmarkPrefix + subDocumentCount);
                    builder.EndBookmark(bookmarkPrefix + subDocumentCount);
                }

                // Iterate through each NUMPAGES field in the section and replace the field with a PAGEREF field referring to the bookmark of the current subdocument
                // This bookmark is positioned at the end of the sub document but does not exist yet. It is inserted when a section with restart page numbering or the last 
                // Section is encountered
                Node[] nodes = section.GetChildNodes(NodeType.FieldStart, true).ToArray();
                foreach (FieldStart fieldStart in nodes)
                {
                    if (fieldStart.FieldType == FieldType.FieldNumPages)
                    {
                        // Get the field code
                        string fieldCode = GetFieldCode(fieldStart);
                        // Since the NUMPAGES field does not take any additional parameters we can assume the remaining part of the field
                        // Code after the fieldname are the switches. We will use these to help recreate the NUMPAGES field as a PAGEREF field
                        string fieldSwitches = fieldCode.Replace(numPagesFieldName, "").Trim();

                        // Inserting the new field directly at the FieldStart node of the original field will cause the new field to
                        // Not pick up the formatting of the original field. To counter this insert the field just before the original field
                        // If a previous run cannot be found then we are forced to use the FieldStart node
                        Node previousNode = fieldStart.PreviousSibling ?? fieldStart;
                        
                        // Insert a PAGEREF field at the same position as the field
                        builder.MoveTo(previousNode);
                        // This will insert a new field with a code like " PAGEREF _SubDocumentEnd0 *\MERGEFORMAT "
                        Field newField = builder.InsertField(string.Format(" {0} {1}{2} {3} ", pageRefFieldName,
                            bookmarkPrefix, subDocumentCount, fieldSwitches));

                        // The field will be inserted before the referenced node. Move the node before the field instead
                        previousNode.ParentNode.InsertBefore(previousNode, newField.Start);

                        // Remove the original NUMPAGES field from the document
                        RemoveField(fieldStart);
                    }
                }
            }
        }
        //ExEnd:ConvertNumPageFieldsToPageRef
        
        //ExStart:GetRemoveField
        private static void RemoveField(FieldStart fieldStart)
        {
            Node currentNode = fieldStart;
            bool isRemoving = true;
            while (currentNode != null && isRemoving)
            {
                if (currentNode.NodeType == NodeType.FieldEnd)
                    isRemoving = false;

                Node nextNode = currentNode.NextPreOrder(currentNode.Document);
                currentNode.Remove();
                currentNode = nextNode;
            }
        }

        private static string GetFieldCode(FieldStart fieldStart)
        {
            StringBuilder builder = new StringBuilder();

            for (Node node = fieldStart;
                node != null && node.NodeType != NodeType.FieldSeparator &&
                node.NodeType != NodeType.FieldEnd;
                node = node.NextPreOrder(node.Document))
            {
                // Use text only of Run nodes to avoid duplication
                if (node.NodeType == NodeType.Run)
                    builder.Append(node.GetText());
            }

            return builder.ToString();
        }
        //ExEnd:GetRemoveField

        [Test]
        public static void DifferentPageSetup()
        {
            //ExStart:DifferentPageSetup
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Set the source document to continue straight after the end of the destination document.
            // If some page setup settings are different then this may not work and the source document will appear 
            // On a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // To ensure this does not happen when the source document has different page setup settings make sure the
            // Settings are identical between the last section of the destination document.
            // If there are further continuous sections that follow on in the source document then this will need to be 
            // Repeated for those sections as well.
            srcDoc.FirstSection.PageSetup.PageWidth = dstDoc.LastSection.PageSetup.PageWidth;
            srcDoc.FirstSection.PageSetup.PageHeight = dstDoc.LastSection.PageSetup.PageHeight;
            srcDoc.FirstSection.PageSetup.Orientation = dstDoc.LastSection.PageSetup.Orientation;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "DifferentPageSetup.docx");
            //ExEnd:DifferentPageSetup
        }

        [Test]
        public static void JoinContinuous()
        {
            //ExStart:JoinContinuous
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Make the document appear straight after the destination documents content
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Append the source document using the original styles found in the source document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinContinuous.docx");
            //ExEnd:JoinContinuous
        }

        [Test]
        public static void JoinNewPage()
        {
            //ExStart:JoinNewPage
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Set the appended document to start on a new page
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

            // Append the source document using the original styles found in the source document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "JoinNewPage.docx");
            //ExEnd:JoinNewPage
        }

        [Test]
        public static void KeepSourceFormatting()
        {
            //ExStart:KeepSourceFormatting
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Keep the formatting from the source document when appending it to the destination document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the joined document to disk
            dstDoc.Save(ArtifactsDir + "KeepSourceFormatting.docx");
            //ExEnd:KeepSourceFormatting
        }

        [Test]
        public static void KeepSourceTogether()
        {
            //ExStart:KeepSourceTogether
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");
            
            // Set the source document to appear straight after the destination document's content
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Iterate through all sections in the source document
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.KeepWithNext = true;
            }

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "KeepSourceTogether.docx");
            //ExEnd:KeepSourceTogether
        }

        [Test]
        public static void LinkHeadersFooters()
        {
            //ExStart:LinkHeadersFooters
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Set the appended document to appear on a new page
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

            // Link the headers and footers in the source document to the previous section
            // This will override any headers or footers already found in the source document
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(true);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "LinkHeadersFooters.docx");
            //ExEnd:LinkHeadersFooters
        }

        [Test]
        public static void ListKeepSourceFormatting()
        {
            //ExStart:ListKeepSourceFormatting
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");

            // Append the content of the document so it flows continuously
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "ListKeepSourceFormatting.docx");
            //ExEnd:ListKeepSourceFormatting
        }

        [Test]
        public static void ListUseDestinationStyles()
        {
            //ExStart:ListUseDestinationStyles
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Document destination with list.docx");

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

        [Test]
        public static void PrependDocument()
        {
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Append the source document to the destination document. This causes the result to have line spacing problems
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Instead prepend the content of the destination document to the start of the source document
            // This results in the same joined document but with no line spacing issues
            DoPrepend(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            dstDoc.Save(ArtifactsDir + "PrependDocument.docx");
        }

        public static void DoPrepend(Document dstDoc, Document srcDoc, ImportFormatMode mode)
        {
            // Loop through all sections in the source document
            // Section nodes are immediate children of the Document node so we can just enumerate the Document
            ArrayList sections = new ArrayList(srcDoc.Sections.ToArray());

            // Reverse the order of the sections so they are prepended to start of the destination document in the correct order
            sections.Reverse();

            foreach (Section srcSection in sections)
            {
                // Import the nodes from the source document
                Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                // Now the new section node can be prepended to the destination document
                // Note how PrependChild is used instead of AppendChild. This is the only line changed compared 
                // To the original method
                dstDoc.PrependChild(dstSection);
            }
        }

        [Test]
        public static void RemoveSourceHeadersFooters()
        {
            //ExStart:RemoveSourceHeadersFooters
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Remove the headers and footers from each of the sections in the source document
            foreach (Section section in srcDoc.Sections)
            {
                section.ClearHeadersFooters();
            }

            // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
            // For HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
            // Document. This should set to false to avoid this behavior
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "RemoveSourceHeadersFooters.docx");
            //ExEnd:RemoveSourceHeadersFooters
        }

        [Test]
        public static void RestartPageNumbering()
        {
            //ExStart:RestartPageNumbering
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Set the appended document to appear on the next page
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            // Restart the page numbering for the document to be appended
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "RestartPageNumbering.docx");
            //ExEnd:RestartPageNumbering
        }

        [Test]
        public static void UnlinkHeadersFooters()
        {
            //ExStart:UnlinkHeadersFooters
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
            // From the destination document
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(false);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            
            dstDoc.Save(ArtifactsDir + "UnlinkHeadersFooters.docx");
            //ExEnd:UnlinkHeadersFooters
        }

        [Test]
        public static void UpdatePageLayout()
        {
            //ExStart:UpdatePageLayout
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document 
            // Is appended then any changes made after will not be reflected in the rendered output
            dstDoc.UpdatePageLayout();

            // Join the documents
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // For the changes to be updated to rendered output, UpdatePageLayout must be called again
            // If not called again the appended document will not appear in the output of the next rendering
            dstDoc.UpdatePageLayout();

            dstDoc.Save(ArtifactsDir + "UpdatePageLayout.docx");
            //ExEnd:UpdatePageLayout
        }

        [Test]
        public static void UseDestinationStyles()
        {
            //ExStart:UseDestinationStyles
            Document srcDoc = new Document(MyDir + "Document source.docx");
            Document dstDoc = new Document(MyDir + "Northwind traders.docx");

            // Append the source document using the styles of the destination document
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            dstDoc.Save(ArtifactsDir + "UseDestinationStyles.docx");
            //ExEnd:UseDestinationStyles
        }

        [Test]
        public static void SmartStyleBehavior()
        {
            //ExStart:SmartStyleBehavior
            Document srcDoc = new Document(MyDir + "Source document.docx");
            Document dstDoc = new Document(MyDir + "Destination document.docx");

            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            ImportFormatOptions options = new ImportFormatOptions();
            options.SmartStyleBehavior = true;
            builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            //ExEnd:SmartStyleBehavior
        }

        [Test]
        public static void KeepSourceNumbering()
        {
            //ExStart:KeepSourceNumbering
            Document srcDoc = new Document(MyDir + "Source document.docx");
            Document dstDoc = new Document(MyDir + "Destination document.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            // Keep source list formatting when importing numbered paragraphs
            importFormatOptions.KeepSourceNumbering = true;
            
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting,
                importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, false);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(ArtifactsDir + "output.docx");
            //ExEnd:KeepSourceNumbering
        }

        [Test]
        public static void IgnoreTextBoxes()
        {
            //ExStart:IgnoreTextBoxes
            Document srcDoc = new Document(MyDir + "Source document.docx");
            Document dstDoc = new Document(MyDir + "Destination document.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            // Keep the source text boxes formatting when importing
            importFormatOptions.IgnoreTextBoxes = false;
            
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting,
                importFormatOptions);

            ParagraphCollection srcParas = srcDoc.FirstSection.Body.Paragraphs;
            foreach (Paragraph srcPara in srcParas)
            {
                Node importedNode = importer.ImportNode(srcPara, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(ArtifactsDir + "output.docx");
            //ExEnd:IgnoreTextBoxes
        }

        [Test]
        public static void IgnoreHeaderFooter()
        {
            // ExStart:IgnoreHeaderFooter
            Document srcDocument = new Document(MyDir + "Source document.docx");
            Document dstDocument = new Document(MyDir + "Destination document.docx");

            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            importFormatOptions.IgnoreHeaderFooter = false;

            dstDocument.AppendDocument(srcDocument, ImportFormatMode.KeepSourceFormatting, importFormatOptions);
            dstDocument.Save(ArtifactsDir + "IgnoreHeaderFooter.docx");
            // ExEnd:IgnoreHeaderFooter
        }
    }
}