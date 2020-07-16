using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class GetFieldNames
    {
        [Test]
        public static void GetMailMergeFieldNames()
        {
            //ExStart:GetFieldNames
            Document doc = new Document();
            // Shows how to get names of all merge fields in a document
            string[] fieldNames = doc.MailMerge.GetFieldNames();
            //ExEnd:GetFieldNames
            Console.WriteLine("\nDocument have " + fieldNames.Length + " fields.");
        }

        [Test]
        public static void MappedDataFields()
        {
            //ExStart:MappedDataFields
            Document doc = new Document();
            // Shows how to add a mapping when a merge field in a document and a data field in a data source have different names
            doc.MailMerge.MappedDataFields.Add("MyFieldName_InDocument", "MyFieldName_InDataSource");
            //ExEnd:MappedDataFields
        }

        [Test]
        public static void DeleteFields()
        {
            //ExStart:DeleteFields
            Document doc = new Document();
            // Shows how to delete all merge fields from a document without executing mail merge
            doc.MailMerge.DeleteFields();
            //ExEnd:DeleteFields
        }
    }
}