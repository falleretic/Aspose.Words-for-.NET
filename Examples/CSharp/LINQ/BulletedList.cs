using Aspose.Words.Reporting;

namespace Aspose.Words.Examples.CSharp.LINQ
{
    class BulletedList : TestDataHelper
    {
        public static void Run()
        {
            CreateBulletedList();
        }

        public static void CreateBulletedList()
        {
            //ExStart:BulletedList
            Document doc = new Document(LinqDir + "BulletedList.doc");

            // Create a Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            // Execute the build report
            engine.BuildReport(doc, Common.GetClients(), "clients");

            // Save the finished document to disk
            doc.Save(ArtifactsDir + "CreateBulletedList.docx");
            //ExEnd:BulletedList
        }
    }
}