using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class InParagraphList : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:InParagraphList
            Document doc = new Document(LinqDir + "InParagraphList.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "InParagraphList.docx");
            //ExEnd:InParagraphList
        }
    }
}