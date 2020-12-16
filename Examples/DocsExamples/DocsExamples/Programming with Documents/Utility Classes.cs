using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class UtilityClasses
    {
        [Test]
        public static void ConvertBetweenMeasurementUnits()
        {
            //ExStart:ConvertBetweenMeasurementUnits
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = ConvertUtil.InchToPoint(1.0);
            pageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0);
            pageSetup.LeftMargin = ConvertUtil.InchToPoint(1.5);
            pageSetup.RightMargin = ConvertUtil.InchToPoint(1.5);
            pageSetup.HeaderDistance = ConvertUtil.InchToPoint(0.2);
            pageSetup.FooterDistance = ConvertUtil.InchToPoint(0.2);
            //ExEnd:ConvertBetweenMeasurementUnits
        }

        [Test]
        public static void UseControlCharacters()
        {
            //ExStart:UseControlCharacters
            const string text = "test\r";
            // Replace "\r" control character with "\r\n".
            string replace = text.Replace(ControlChar.Cr, ControlChar.CrLf);
            //ExEnd:UseControlCharacters
        }
    }
}