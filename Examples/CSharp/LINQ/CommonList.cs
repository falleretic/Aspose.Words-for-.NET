using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class CommonList : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:CommonList
            Document doc = new Document(LinqDir + "CommonList.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "CommonList.docx");
            //ExEnd:CommonList
        }
    }
}