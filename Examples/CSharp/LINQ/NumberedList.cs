using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class NumberedList : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:NumberedList
            Document doc = new Document(LinqDir + "NumberedList.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "NumberedList.docx");
            //ExEnd:NumberedList
        }
    }
}