using System.IO;
using System.Linq;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Split_Documents
{
    class SplitDocument : TestDataHelper
    {
        [Test]
        public static void SplitDocumentByHeadingsHtml()
        {
            //ExStart:SplitDocumentByHeadingsHtml
            // Open a Word document
            Document doc = new Document(MyDir + "Rendering.docx");
 
            HtmlSaveOptions options = new HtmlSaveOptions();
            // Split a document into smaller parts, in this instance split by heading
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
 
            // Save the output file
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SplitDocumentByHeadings.html", options);
            //ExEnd:SplitDocumentByHeadingsHtml
        }

        [Test]
        public static void SplitDocumentBySectionsHtml()
        {
            // Open a Word document
            Document doc = new Document(MyDir + "Rendering.docx");
 
            //ExStart:SplitDocumentBySectionsHtml
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;
            //ExEnd:SplitDocumentBySectionsHtml
            
            // Save the output file
            doc.Save(ArtifactsDir + "HtmlSaveOptionsEx.SplitDocumentBySections.html", options);
        }

        [Test]
        public static void SplitDocumentBySections()
        {
            //ExStart:SplitDocumentBySections
            // Open a Word document
            Document doc = new Document(MyDir + "Big document.docx");

            for (int i = 0; i < doc.Sections.Count; i++)
            {
                // Split a document into smaller parts, in this instance split by section
                Section section = doc.Sections[i].Clone();

                Document newDoc = new Document();
                newDoc.Sections.Clear();

                Section newSection = (Section) newDoc.ImportNode(section, true);
                newDoc.Sections.Add(newSection);

                // Save each section as a separate document
                newDoc.Save(ArtifactsDir + $"SplitDocumentBySectionsOut_{i}.docx");
            }
            //ExEnd:SplitDocumentBySections
        }

        [Test]
        public static void SplitDocumentPageByPage()
        {
            //ExStart:SplitDocumentPageByPage
            // Open a Word document
            Document doc = new Document(MyDir + "Big document.docx");

            // Split nodes in the document into separate pages
            DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

            // Save each page as a separate document
            for (int page = 1; page <= doc.PageCount; page++)
            {
                Document pageDoc = splitter.GetDocumentOfPage(page);
                pageDoc.Save(ArtifactsDir + $"SplitDocumentPageByPageOut_{page}.docx");
            }
            //ExEnd:SplitDocumentPageByPage

            MergeDocuments();
        }

        [Test]
        //ExStart:MergeSplitDocuments
        public static void MergeDocuments()
        {
            // Find documents using for merge
            FileSystemInfo[] documentPaths = new DirectoryInfo(ArtifactsDir)
                .GetFileSystemInfos("SplitDocumentPageByPageOut_*.docx").OrderBy(f => f.CreationTime).ToArray();
            string sourceDocumentPath =
                Directory.GetFiles(ArtifactsDir, "SplitDocumentPageByPageOut_1.docx", SearchOption.TopDirectoryOnly)[0];

            // Open the first part of the resulting document
            Document sourceDoc = new Document(sourceDocumentPath);

            // Create a new resulting document
            Document mergedDoc = new Document();
            DocumentBuilder mergedDocBuilder = new DocumentBuilder(mergedDoc);

            // Merge document parts one by one
            foreach (FileSystemInfo documentPath in documentPaths)
            {
                if (documentPath.FullName == sourceDocumentPath)
                    continue;

                mergedDocBuilder.MoveToDocumentEnd();
                mergedDocBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);
                sourceDoc = new Document(documentPath.FullName);
            }

            // Save the output file
            mergedDoc.Save(ArtifactsDir + "MergeDocuments.docx");
        }
        //ExEnd:MergeSplitDocuments

        [Test]
        public static void SplitDocumentByPageRange()
        {
            //ExStart:SplitDocumentByPageRange
            // Open a Word document
            Document doc = new Document(MyDir + "Big document.docx");
 
            // Split nodes in the document into separate pages
            DocumentPageSplitter splitter = new DocumentPageSplitter(doc);
 
            // Get part of the document
            Document pageDoc = splitter.GetDocumentOfPageRange(3,6);
            pageDoc.Save(ArtifactsDir + "SplitDocumentByPageRangeOut.docx");
            //ExEnd:SplitDocumentByPageRange
        }
    }
}

