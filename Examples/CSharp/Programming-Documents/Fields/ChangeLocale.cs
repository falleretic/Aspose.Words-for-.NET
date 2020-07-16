using System;
using System.Globalization;
using System.Threading;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class ChangeLocale : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:ChangeLocale
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Date");

            // Store the current culture so it can be set back once mail merge is complete
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            // Set to German language so dates and numbers are formatted using this culture during mail merge
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            // Execute mail merge
            doc.MailMerge.Execute(new[] { "Date" }, new object[] { DateTime.Now });
            
            // Restore the original culture
            Thread.CurrentThread.CurrentCulture = currentCulture;
            
            doc.Save(ArtifactsDir + "Field.ChangeLocale.doc");
            //ExEnd:ChangeLocale
        }
    }
}