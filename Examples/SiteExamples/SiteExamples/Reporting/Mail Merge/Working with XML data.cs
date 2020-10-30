﻿using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace SiteExamples.Reporting.Mail_Merge
{
    class WorkingWithXMLData : SiteExamplesBase
    {
        [Test]
        public static void XmlMailMerge()
        {
            //ExStart:XmlMailMerge
            DataSet customersDs = new DataSet();
            customersDs.ReadXml(MyDir + "Mail merge data - Customers.xml");

            Document doc = new Document(MyDir + "Mail merge destinations - Registration complete.docx");
            // Execute mail merge to fill the template with data from XML using DataTable
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            doc.Save(ArtifactsDir + "XmlMailMerge.docx");
            //ExEnd:XmlMailMerge
        }

        [Test]
        public static void NestedMailMerge()
        {
            //ExStart:NestedMailMerge
            DataSet pizzaDs = new DataSet();

            // Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml
            // To see examples of how to set up relations manually check the corresponding documentation of this sample
            pizzaDs.ReadXml(MyDir + "Mail merge data - Orders.xml");

            Document doc = new Document(MyDir + "Mail merge destinations - Invoice.docx");

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // Execute the nested mail merge with regions
            doc.MailMerge.ExecuteWithRegions(pizzaDs);

            doc.Save(ArtifactsDir + "MailMerge.NestedMailMerge.docx");
            //ExEnd:NestedMailMerge
        }

        [Test]
        public static void MustacheSyntax()
        {
            //ExStart:MailMergeUsingMustacheSyntax
            DataSet ds = new DataSet();
            ds.ReadXml(MyDir + "Mail merge data - Vendors.xml");

            // Open a template document
            Document doc = new Document(MyDir + "Mail merge destinations - Vendor.docx");

            doc.MailMerge.UseNonMergeFields = true;

            // Execute mail merge to fill the template with data from XML using DataSet
            doc.MailMerge.ExecuteWithRegions(ds);
            
            doc.Save(ArtifactsDir + "MailMerge.UsingMustacheSyntax.docx");
            //ExEnd:MailMergeUsingMustacheSyntax
        }

        [Test]
        public static void LINQtoXmlMailMerge()
        {
            XElement orderXml = XElement.Load(MyDir + "Mail merge data - Purchase order.xml");

            // Query the purchase order xml file using LINQ to extract the order items 
            // Into an object of an anonymous type. 
            //
            // Make sure you give the properties of the anonymous type the same names as 
            // The MERGEFIELD fields in the document.
            //
            // To pass the actual values stored in the XML element or attribute to Aspose.Words, 
            // we need to cast them to string. This is to prevent the XML tags being inserted into the final document when
            // The XElement or XAttribute objects are passed to Aspose.Words.
            //ExStart:LINQtoXMLMailMergeorderItems
            var orderItems =
                from order in orderXml.Descendants("Item")
                select new
                {
                    PartNumber = (string) order.Attribute("PartNumber"),
                    ProductName = (string) order.Element("ProductName"),
                    Quantity = (string) order.Element("Quantity"),
                    USPrice = (string) order.Element("USPrice"),
                    Comment = (string) order.Element("Comment"),
                    ShipDate = (string) order.Element("ShipDate")
                };
            //ExEnd:LINQtoXMLMailMergeorderItems
            
            //ExStart:LINQToXMLQueryForDeliveryAddress
            var deliveryAddress =
                from delivery in orderXml.Elements("Address")
                where ((string) delivery.Attribute("Type") == "Shipping")
                select new
                {
                    Name = (string) delivery.Element("Name"),
                    Country = (string) delivery.Element("Country"),
                    Zip = (string) delivery.Element("Zip"),
                    State = (string) delivery.Element("State"),
                    City = (string) delivery.Element("City"),
                    Street = (string) delivery.Element("Street")
                };
            //ExEnd:LINQToXMLQueryForDeliveryAddress

            // Create custom Aspose.Words mail merge data sources based on the LINQ queries
            MyMailMergeDataSource orderItemsDataSource = new MyMailMergeDataSource(orderItems, "Items");
            MyMailMergeDataSource deliveryDataSource = new MyMailMergeDataSource(deliveryAddress);
            //ExStart:LINQToXMLMailMerge
            Document doc = new Document(MyDir + "Mail merge destinations - LINQ.docx");

            // Fill the document with data from our data sources
            // Using mail merge regions for populating the order items table is required
            // Because it allows the region to be repeated in the document for each order item
            doc.MailMerge.ExecuteWithRegions(orderItemsDataSource);

            // The standard mail merge without regions is used for the delivery address
            doc.MailMerge.Execute(deliveryDataSource);

            doc.Save(ArtifactsDir + "MailMerge.LINQtoXML.docx");
            //ExEnd:LINQToXMLMailMerge
        }

        /// <summary>
        /// Aspose.Words does not accept LINQ queries as an input for mail merge directly, 
        /// But provides a generic mechanism which allows mail merges from any data source.
        /// 
        /// This class is a simple implementation of the Aspose.Words custom mail merge data source 
        /// Interface that accepts a LINQ query (in fact any IEnumerable object).
        /// Aspose.Words calls this class during the mail merge to retrieve the data.
        /// </summary>
        //ExStart:MyMailMergeDataSource 
        public class MyMailMergeDataSource : IMailMergeDataSource
        //ExEnd:MyMailMergeDataSource 
        {
            /// <summary>
            /// Creates a new instance of a custom mail merge data source.
            /// </summary>
            /// <param name="data">Data returned from a LINQ query.</param>
            //ExStart:MyMailMergeDataSourceConstructor 
            public MyMailMergeDataSource(IEnumerable data)
            {
                mEnumerator = data.GetEnumerator();
            }
            //ExEnd:MyMailMergeDataSourceConstructor

            /// <summary>
            /// Creates a new instance of a custom mail merge data source, for mail merge with regions.
            /// </summary>
            /// <param name="data">Data returned from a LINQ query.</param>
            /// <param name="tableName">Name of the data source is only used when you perform mail merge with regions. 
            /// If you prefer to use the simple mail merge then use constructor with one parameter.</param>          
            //ExStart:MyMailMergeDataSourceConstructorWithDataTable
            public MyMailMergeDataSource(IEnumerable data, string tableName)
            {
                mEnumerator = data.GetEnumerator();
                TableName = tableName;
            }
            //ExEnd:MyMailMergeDataSourceConstructorWithDataTable

            /// <summary>
            /// Aspose.Words calls this method to get a value for every data field.
            /// 
            /// This is a simple "generic" implementation of a data source that can work over 
            /// Any IEnumerable collection. This implementation assumes that the merge field
            /// Name in the document matches the name of a public property on the object
            /// In the collection and uses reflection to get the value of the property.
            /// </summary>
            //ExStart:MyMailMergeDataSourceGetValue
            public bool GetValue(string fieldName, out object fieldValue)
            {
                // Use reflection to get the property by name from the current object
                object obj = mEnumerator.Current;

                Type currentRecordType = obj.GetType();
                PropertyInfo property = currentRecordType.GetProperty(fieldName);
                if (property != null)
                {
                    fieldValue = property.GetValue(obj, null);
                    return true;
                }

                // Return False to the Aspose.Words mail merge engine to indicate the field was not found
                fieldValue = null;
                return false;
            }
            //ExEnd:MyMailMergeDataSourceGetValue

            /// <summary>
            /// Moves to the next record in the collection.
            /// </summary>            
            //ExStart:MyMailMergeDataSourceMoveNext
            public bool MoveNext()
            {
                return mEnumerator.MoveNext();
            }
            //ExEnd:MyMailMergeDataSourceMoveNext

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            //ExStart:MyMailMergeDataSourceTableName
            public string TableName { get; }
            //ExEnd:MyMailMergeDataSourceTableName

            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }

            private readonly IEnumerator mEnumerator;
        }
    }
}