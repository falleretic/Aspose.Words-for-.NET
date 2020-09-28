using System.IO;
using System.Text;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class WorkingWithHtmlLoadOptions : TestDataHelper
    {
        [Test]
        public static void PreferredControlType()
        {
            //ExStart:LoadAndSaveHtmlFormFieldasContentControlinDOCX
            const string html = @"
                <html>
                    <select name='ComboBox' size='1'>
                        <option value='val1'>item1</option>
                        <option value='val2'></option>                        
                    </select>
                </html>
            ";
            
            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
            htmlLoadOptions.PreferredControlType = HtmlControlType.StructuredDocumentTag;

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
            doc.Save(ArtifactsDir + "HtmlLoadOptionsEx.PreferredControlType.docx", SaveFormat.Docx);
            //ExEnd:LoadAndSaveHtmlFormFieldasContentControlinDOCX
        }
    }
}