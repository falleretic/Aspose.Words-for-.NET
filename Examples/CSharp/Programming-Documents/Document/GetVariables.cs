using System;
using System.Collections.Generic;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class GetVariables : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:GetVariables
            Document doc = new Document(DocumentDir + "TestFile.doc");
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
    }
}