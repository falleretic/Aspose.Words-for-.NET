using System.Data;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class NestedMailMerge : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:NestedMailMerge
            DataSet pizzaDs = new DataSet();

            // Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml
            // To see examples of how to set up relations manually check the corresponding documentation of this sample
            pizzaDs.ReadXml(MailMergeDir + "CustomerData.xml");

            Document doc = new Document(MailMergeDir + "Invoice Template.doc");

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // Execute the nested mail merge with regions
            doc.MailMerge.ExecuteWithRegions(pizzaDs);

            doc.Save(ArtifactsDir + "NestedMailMerge.docx");
            //ExEnd:NestedMailMerge
            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0, "There was a problem with mail merge");
        }
    }
}