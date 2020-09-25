using System;
using System.IO;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class ApplyLicense : TestDataHelper
    {
        [Test]
        public static void ApplyLicenseFromFile()
        {
            //ExStart:ApplyLicenseFromFile
            License license = new License();

            // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
            // You can also use the additional overload to load a license from a stream, this is useful for instance when the 
            // license is stored as an embedded resource
            try
            {
                license.SetLicense("Aspose.Words.lic");
                Console.WriteLine("License set successfully.");
            }
            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
            //ExEnd:ApplyLicenseFromFile
        }

        [Test]
        public static void ApplyLicenseFromStream()
        {
            //ExStart:ApplyLicenseFromStream
            License license = new License();

            try
            {
                // Initializes a license from a stream 
                MemoryStream stream = new MemoryStream(File.ReadAllBytes(@"Aspose.Words.lic"));
                license.SetLicense(stream);
                Console.WriteLine("License set successfully.");
            }
            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }
            //ExEnd:ApplyLicenseFromStream
        }

        [Test]
        public static void ApplyMeteredLicense()
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