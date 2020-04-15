using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class CommonMasterDetail : TestDataHelper
    {
        [Test]
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