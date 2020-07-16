using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ApplyMeteredLicense : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ApplyMeteredLicense
            try
            {
                // Set metered public and private keys
                Metered metered = new Metered();
                // Access the setMeteredKey property and pass public and private keys as parameters
                metered.SetMeteredKey("*****", "*****");

                // Load the document from disk
                Document doc = new Document(QuickStartDir + "Template.doc");

                // Get the page count of document
                Console.WriteLine(doc.PageCount);
            }
            catch (Exception e)
            {
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
            //ExEnd:ApplyMeteredLicense
        }
    }
}