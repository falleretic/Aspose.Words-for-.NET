using System.Diagnostics;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class AutoFitTableToWindow : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // ExStart:AutoFitTableToPageWidth
            Document doc = new Document(TablesDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Autofit the first table to the page width
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            doc.Save(ArtifactsDir + "AutoFitTableToWindow.docx");

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Percent,
                "PreferredWidth type is not percent");
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 100,
                "PreferredWidth value is different than 100");
            //ExEnd:AutoFitTableToPageWidth
        }
    }
}