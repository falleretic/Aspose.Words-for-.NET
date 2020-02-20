using System;
using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BulletedList
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_LINQ();

            CreateBulletedList(dataDir);
        }

        public static void CreateBulletedList(string dataDir)
        {
            // ExStart:BulletedList
            // Load the template document
            Document doc = new Document(dataDir + "BulletedList.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();

            // Execute the build report
            engine.BuildReport(doc, Common.GetClients(), "clients");

            // Save the finished document to disk
            doc.Save(dataDir + RunExamples.GetOutputFilePath("BulletedList.doc"));
            // ExEnd:BulletedList

            Console.WriteLine(
                "\nBulleted list template document is populated with the data about clients.\nFile saved at " +
                dataDir);
        }
    }
}