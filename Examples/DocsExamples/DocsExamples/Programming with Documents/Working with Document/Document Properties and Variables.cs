﻿using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Properties;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Document
{
    internal class DocumentPropertiesAndVariables : DocsExamplesBase
    {
        [Test]
        public static void GetVariables()
        {
            //ExStart:GetVariables
            Document doc = new Document(MyDir + "Document.docx");
            
            string variables = "";
            foreach (KeyValuePair<string, string> entry in doc.Variables)
            {
                string name = entry.Key;
                string value = entry.Value;
                if (variables == "")
                {
                    variables = "Name: " + name + "," + "Value: {1}" + value;
                }
                else
                {
                    variables = variables + "Name: " + name + "," + "Value: {1}" + value;
                }
            }
            //ExEnd:GetVariables

            Console.WriteLine("\nDocument have following variables " + variables);
        }

        [Test]
        public static void EnumerateProperties()
        {
            //ExStart:EnumerateProperties            
            Document doc = new Document(MyDir + "Properties.docx");
            
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
        public static void AddCustomDocumentProperties()
        {
            //ExStart:AddCustomDocumentProperties            
            Document doc = new Document(MyDir + "Properties.docx");

            CustomDocumentProperties customDocumentProperties = doc.CustomDocumentProperties;
            
            if (customDocumentProperties["Authorized"] != null) return;
            
            customDocumentProperties.Add("Authorized", true);
            customDocumentProperties.Add("Authorized By", "John Smith");
            customDocumentProperties.Add("Authorized Date", DateTime.Today);
            customDocumentProperties.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
            customDocumentProperties.Add("Authorized Amount", 123.45);
            //ExEnd:AddCustomDocumentProperties
        }

        [Test]
        public static void RemoveCustomDocumentProperties()
        {
            //ExStart:CustomRemove            
            Document doc = new Document(MyDir + "Properties.docx");
            doc.CustomDocumentProperties.Remove("Authorized Date");
            //ExEnd:CustomRemove
        }

        [Test]
        public static void RemovePersonalInformation()
        {
            //ExStart:RemovePersonalInformation            
            Document doc = new Document(MyDir + "Properties.docx") { RemovePersonalInformation = true };

            doc.Save(ArtifactsDir + "DocumentPropertiesAndVariables.RemovePersonalInformation.docx");
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

            // Retrieve a list of all custom document properties from the file.
            CustomDocumentProperties customProperties = doc.CustomDocumentProperties;
            // Add linked to content property.
            DocumentProperty customProperty = customProperties.AddLinkToContent("Bookmark", "MyBookmark");
            customProperty = customProperties["Bookmark"];

            bool isLinkedToContent = customProperty.IsLinkToContent;
            
            string linkSource = customProperty.LinkSource;
            
            string customPropertyValue = customProperty.Value.ToString();
            //ExEnd:ConfiguringLinkToContent
        }
    }
}