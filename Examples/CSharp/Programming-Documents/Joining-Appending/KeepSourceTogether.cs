using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class KeepSourceTogether : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:KeepSourceTogether
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Document destination with list.docx");
            
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
    }
}