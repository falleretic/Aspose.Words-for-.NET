using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class MulticoloredNumberedList : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:MulticoloredNumberedList
            Document doc = new Document(LinqDir + "MulticoloredNumberedList.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "MulticoloredNumberedList.doc");
            //ExEnd:MulticoloredNumberedList
        }
    }
}