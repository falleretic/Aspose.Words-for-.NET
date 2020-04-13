using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class CommonMasterDetail : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:CommonMasterDetail
            Document doc = new Document(LinqDir + "CommonMasterDetail.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "CommonMasterDetail.docx");
            //ExEnd:CommonMasterDetail
        }
    }
}