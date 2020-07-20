using System.Data;
using System.Diagnostics;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class NestedMailMerge : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:NestedMailMerge
            DataSet pizzaDs = new DataSet();

            // Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml
            // To see examples of how to set up relations manually check the corresponding documentation of this sample
            pizzaDs.ReadXml(MailMergeDir + "Mail merge data - CustomerData.xml");

            Document doc = new Document(MailMergeDir + "Mail merge destinations - Invoice.docx");

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // Execute the nested mail merge with regions
            doc.MailMerge.ExecuteWithRegions(pizzaDs);

            doc.Save(ArtifactsDir + "MailMerge.NestedMailMerge.docx");
            //ExEnd:NestedMailMerge
            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0, "There was a problem with mail merge");
        }
    }
}