namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_and_Appending
{
    class DifferentPageSetup : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:DifferentPageSetup
            Document dstDoc = new Document(JoiningAppendingDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(JoiningAppendingDir + "TestFile.SourcePageSetup.doc");

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
    }
}