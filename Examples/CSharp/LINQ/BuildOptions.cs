using Aspose.Words.Reporting;
using System;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BuildOptions
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_LINQ();

            RemoveEmptyParagraphs(dataDir);
        }

        public static void RemoveEmptyParagraphs(string dataDir)
        {
            // ExStart:RemoveEmptyParagraphs
            // Load the template document
            Document doc = new Document(dataDir + "template_cleanup.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            // Save the finished document to disk
            doc.Save(dataDir + RunExamples.GetOutputFilePath("template_cleanup.docx"));
            // ExEnd:RemoveEmptyParagraphs

            Console.WriteLine(
                "\nEmpty paragraphs are removed from the document successfully.\nFile saved at " + dataDir);
        }
    }
}