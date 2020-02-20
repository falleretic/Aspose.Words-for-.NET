using Aspose.Words.Reporting;
using System;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class SettingBackgroundColor
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LINQ();
            string fileName = "SettingBackgroundColor.docx";
            
            // ExStart:SettingBackgroundColor
            // Load the template document.
            Document doc = new Document(dataDir + fileName);

            // Create a Reporting Engine.
            ReportingEngine engine = new ReportingEngine();

            // Execute the build report.
            engine.BuildReport(doc, new object());

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

            // Save the finished document to disk.
            doc.Save(dataDir);
            // ExEnd:SettingBackgroundColor
            
            Console.WriteLine("\nSet the background color of text and shape successfully.\nFile saved at " + dataDir);
        }
    }
}