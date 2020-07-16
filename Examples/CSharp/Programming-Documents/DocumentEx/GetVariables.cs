using System;
using System.Collections.Generic;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class GetVariables : TestDataHelper
    {
        [Test]
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