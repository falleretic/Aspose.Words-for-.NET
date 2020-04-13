using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class SettingBackgroundColor : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:SettingBackgroundColor
            Document doc = new Document(LinqDir + "SettingBackgroundColor.docx");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, new object());

            doc.Save(ArtifactsDir + "SettingBackgroundColor.docx");
            //ExEnd:SettingBackgroundColor
        }
    }
}