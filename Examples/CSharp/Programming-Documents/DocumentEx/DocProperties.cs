﻿using System;
using Aspose.Words.Properties;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class DocProperties : TestDataHelper
    {
        [Test]
        public static void EnumerateProperties()
        {
            //ExStart:EnumerateProperties            
            Document doc = new Document(DocumentDir + "Properties.docx");
            Console.WriteLine("1. Document name: {0}", doc.OriginalFileName);

            Console.WriteLine("2. Built-in Properties");
            foreach (DocumentProperty prop in doc.BuiltInDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);

            Console.WriteLine("3. Custom Properties");
            foreach (DocumentProperty prop in doc.CustomDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);
            //ExEnd:EnumerateProperties
        }

        [Test]
        public static void CustomAdd()
        {
            //ExStart:CustomAdd            
            Document doc = new Document(DocumentDir + "Properties.docx");

            CustomDocumentProperties props = doc.CustomDocumentProperties;
            if (props["Authorized"] != null) return;
            props.Add("Authorized", true);
            props.Add("Authorized By", "John Smith");
            props.Add("Authorized Date", DateTime.Today);
            props.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
            props.Add("Authorized Amount", 123.45);
            //ExEnd:CustomAdd
        }

        [Test]
        public static void CustomRemove()
        {
            //ExStart:CustomRemove            
            Document doc = new Document(DocumentDir + "Properties.docx");
            doc.CustomDocumentProperties.Remove("Authorized Date");
            //ExEnd:CustomRemove
        }

        [Test]
        public static void RemovePersonalInformation()
        {
            //ExStart:RemovePersonalInformation            
            Document doc = new Document(DocumentDir + "Properties.docx");
            doc.RemovePersonalInformation = true;

            doc.Save(ArtifactsDir + "RemovePersonalInformation.docx");
            //ExEnd:RemovePersonalInformation
        }

        [Test]
        public static void ConfiguringLinkToContent()
        {
            //ExStart:ConfiguringLinkToContent            
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.StartBookmark("MyBookmark");
            builder.Writeln("Text inside a bookmark.");
            builder.EndBookmark("MyBookmark");

            // Retrieve a list of all custom document properties from the file
            CustomDocumentProperties customProperties = doc.CustomDocumentProperties;
            // Add linked to content property
            DocumentProperty customProperty = customProperties.AddLinkToContent("Bookmark", "MyBookmark");
            // Also, accessing the custom document property can be performed by using the property name
            customProperty = customProperties["Bookmark"];

            // Check whether the property is linked to content
            bool isLinkedToContent = customProperty.IsLinkToContent;
            // Get the source of the property
            string source = customProperty.LinkSource;
            // Get the value of the property
            string value = customProperty.Value.ToString();
            //ExEnd:ConfiguringLinkToContent
        }
    }
}