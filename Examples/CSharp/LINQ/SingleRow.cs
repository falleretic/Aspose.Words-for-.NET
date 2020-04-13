using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class SingleRow : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:SingleRow
            Document doc = new Document(LinqDir + "SingleRow.doc");

            // Load the photo and read all bytes
            byte[] imgdata = System.IO.File.ReadAllBytes(LinqDir + "photo.png");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetManager(), "manager");

            doc.Save(ArtifactsDir + "SingleRow.docx");
            //ExEnd:SingleRow
        }
    }
}