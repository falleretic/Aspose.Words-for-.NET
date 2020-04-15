using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkingWithFootnote : TestDataHelper
    {
        [Test]
        public static void SetFootNoteColumns()
        {
            //ExStart:SetFootNoteColumns
            Document doc = new Document(DocumentDir + "TestFile.docx");

            // Specify the number of columns with which the footnotes area is formatted
            doc.FootnoteOptions.Columns = 3;
            
            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:SetFootNoteColumns
        }

        [Test]
        public static void SetFootnoteAndEndNotePosition()
        {
            //ExStart:SetFootnoteAndEndNotePosition
            Document doc = new Document(DocumentDir + "TestFile.docx");

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
            Document doc = new Document(DocumentDir + "TestFile.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Some text");
            builder.InsertFootnote(FootnoteType.Endnote, "Eootnote text.");

            EndnoteOptions option = doc.EndnoteOptions;
            option.RestartRule = FootnoteNumberingRule.RestartPage;
            option.Position = EndnotePosition.EndOfSection;

            doc.Save(ArtifactsDir + "TestFile.doc");
            //ExEnd:SetEndnoteOptions
        }
    }
}