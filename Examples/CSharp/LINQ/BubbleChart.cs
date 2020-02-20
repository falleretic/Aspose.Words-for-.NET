using System;
using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BubbleChart
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_LINQ();

            CreateBubbleChart(dataDir);
        }

        public static void CreateBubbleChart(string dataDir)
        {
            // ExStart:BubbleChart
            // Load the template document
            Document doc = new Document(dataDir + "BubbleChart.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();

            // Execute the build report
            engine.BuildReport(doc, Common.GetContracts(), "contracts");

            // Save the finished document to disk
            doc.Save(dataDir + RunExamples.GetOutputFilePath("BubbleChart.docx"));
            // ExEnd:BubbleChart

            Console.WriteLine(
                "\nBubble chart template document is populated with the data about contracts.\nFile saved at " +
                dataDir);
        }
    }
}