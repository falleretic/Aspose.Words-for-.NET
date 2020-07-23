using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class UpdatePageLayout : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:UpdatePageLayout
            Document srcDoc = new Document(JoiningAppendingDir + "Document source.docx");
            Document dstDoc = new Document(JoiningAppendingDir + "Northwind traders.docx");

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
    }
}