using Aspose.Words.Reporting;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.LINQ_Reporting_Engine
{
    internal class Lists : TestDataHelper
    {
        [Test]
        public static void CreateBulletedList()
        {
            //ExStart:BulletedList
            Document doc = new Document(LinqDir + "Reporting engine template - Bulleted list.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "ReportingEngine.CreateBulletedList.docx");
            //ExEnd:BulletedList
        }

        [Test]
        public static void CommonList()
        {
            //ExStart:CommonList
            Document doc = new Document(LinqDir + "Reporting engine template - Common list.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.CommonList.docx");
            //ExEnd:CommonList
        }

        [Test]
        public static void InParagraphList()
        {
            //ExStart:InParagraphList
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("<<foreach [in clients]>><<[IndexOf() !=0 ? ”, ”:  ””]>><<[Name]>><</foreach>>");
            
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "ReportingEngine.InParagraphList.docx");
            //ExEnd:InParagraphList
        }

        [Test]
        public static void InTableList()
        {
            //ExStart:InTableList
            Document doc = new Document(LinqDir + "Reporting engine template - Table list.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.InTableList.docx");
            //ExEnd:InTableList
        }

        [Test]
        public static void MulticoloredNumberedList()
        {
            //ExStart:MulticoloredNumberedList
            Document doc = new Document(LinqDir + "Reporting engine template - Multicolored numbered list.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "ReportingEngine.MulticoloredNumberedList.doc");
            //ExEnd:MulticoloredNumberedList
        }

        [Test]
        public static void NumberedList()
        {
            //ExStart:NumberedList
            Document doc = new Document(LinqDir + "Reporting engine template - Numbered list.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetClients(), "clients");

            doc.Save(ArtifactsDir + "ReportingEngine.NumberedList.docx");
            //ExEnd:NumberedList
        }
    }
}