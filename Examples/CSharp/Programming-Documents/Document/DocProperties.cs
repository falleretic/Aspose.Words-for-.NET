using System;
using Aspose.Words.Properties;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocProperties : TestDataHelper
    {
        public static void Run()
        {
            // Enumerates through all built-in and custom properties in a document
            EnumerateProperties();
            // Checks if a custom property with a given name exists in a document and adds few more custom document properties
            CustomAdd();
            // Removes a custom document property
            CustomRemove();

            RemovePersonalInformation();
            ConfiguringLinkToContent();
        }

        public static void EnumerateProperties()
        {
            //ExStart:EnumerateProperties            
            Document doc = new Document(DocumentDir + "Properties.doc");
            Console.WriteLine("1. Document name: {0}", doc.OriginalFileName);

            Console.WriteLine("2. Built-in Properties");
            foreach (DocumentProperty prop in doc.BuiltInDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);

            Console.WriteLine("3. Custom Properties");
            foreach (DocumentProperty prop in doc.CustomDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);
            //ExEnd:EnumerateProperties
        }

        public static void CustomAdd()
        {
            //ExStart:CustomAdd            
            Document doc = new Document(DocumentDir + "Properties.doc");

            CustomDocumentProperties props = doc.CustomDocumentProperties;
            if (props["Authorized"] != null) return;
            props.Add("Authorized", true);
            props.Add("Authorized By", "John Smith");
            props.Add("Authorized Date", DateTime.Today);
            props.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
            props.Add("Authorized Amount", 123.45);
            //ExEnd:CustomAdd
        }

        public static void CustomRemove()
        {
            //ExStart:CustomRemove            
            Document doc = new Document(DocumentDir + "Properties.doc");
            doc.CustomDocumentProperties.Remove("Authorized Date");
            //ExEnd:CustomRemove
        }

        public static void RemovePersonalInformation()
        {
            //ExStart:RemovePersonalInformation            
            Document doc = new Document(DocumentDir + "Properties.doc");
            doc.RemovePersonalInformation = true;

            doc.Save(ArtifactsDir + "RemovePersonalInformation.docx");
            //ExEnd:RemovePersonalInformation
        }

        public static void ConfiguringLinkToContent()
        {
            //ExStart:ConfiguringLinkToContent            
            Document doc = new Document(DocumentDir + "test.docx");

            // Retrieve a list of all custom document properties from the file
            CustomDocumentProperties customProperties = doc.CustomDocumentProperties;

            // Add linked to content property
            DocumentProperty customProperty = customProperties.AddLinkToContent("PropertyName", "BookmarkName");

            // Also, accessing the custom document property can be performed by using the property name
            customProperty = customProperties["PropertyName"];

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