using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractTextOnly
    {
        [Test]
        public static void Run()
        {
            //ExStart:ExtractTextOnly
            Document doc = new Document();

            // Enter a dummy field into the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Field");

            // GetText will retrieve all field codes and special characters
            Console.WriteLine("GetText() Result: " + doc.GetText());

            // ToString will export the node to the specified format
            // When converted to text it will not retrieve fields code or special characters,
            // but will still contain some natural formatting characters such as paragraph markers etc. 
            // This is the same as "viewing" the document as if it was opened in a text editor
            Console.WriteLine("ToString() Result: " + doc.ToString(SaveFormat.Text));
            //ExEnd:ExtractTextOnly            
        }
    }
}