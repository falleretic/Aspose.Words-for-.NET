using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Document_Content
{
    class WorkingWithFootnotes : DocsExamplesBase
    {
        [Test]
        public static void SetFootNoteColumns()
        {
            //ExStart:SetFootNoteColumns
            Document doc = new Document(MyDir + "Document.docx");

            // Specify the number of columns with which the footnotes area is formatted
            doc.FootnoteOptions.Columns = 3;
            
            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:SetFootNoteColumns
        }

        [Test]
        public static void SetFootnoteAndEndNotePosition()
        {
            //ExStart:SetFootnoteAndEndNotePosition
            Document doc = new Document(MyDir + "Document.docx");

            // Set footnote and endnode position
            doc.FootnoteOptions.Position = FootnotePosition.BeneathText;
            doc.EndnoteOptions.Position = EndnotePosition.EndOfSection;
            
            doc.Save(ArtifactsDir + "TestFile_Out.doc");
            //ExEnd:SetFootnoteAndEndNotePosition
        }

        [Test]
        public static void SetEndnoteOptions()
        {
            //ExStart:SetEndnoteOptions
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Some text");
            builder.InsertFootnote(FootnoteType.Endnote, "Footnote text.");

            EndnoteOptions option = doc.EndnoteOptions;
            option.RestartRule = FootnoteNumberingRule.RestartPage;
            option.Position = EndnotePosition.EndOfSection;

            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:SetEndnoteOptions
        }
    }
}