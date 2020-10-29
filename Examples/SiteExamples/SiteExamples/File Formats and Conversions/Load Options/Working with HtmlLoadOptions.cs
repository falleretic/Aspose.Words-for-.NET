using System.IO;
using System.Text;
using Aspose.Words;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Load_Options
{
    internal class WorkingWithHtmlLoadOptions : SiteExamplesBase
    {
        [Test, Description("Shows how to set the preferred type for the imported HTML elements.")]
        public void PreferredControlType()
        {
            //ExStart:LoadHtmlElementsWithPreferredControlType
            const string html = @"
                <html>
                    <select name='ComboBox' size='1'>
                        <option value='val1'>item1</option>
                        <option value='val2'></option>                        
                    </select>
                </html>
            ";
            
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.PreferredControlType = HtmlControlType.StructuredDocumentTag;

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), loadOptions);
            doc.Save(ArtifactsDir + "HtmlLoadOptions.PreferredControlType.docx", SaveFormat.Docx);
            //ExEnd:LoadHtmlElementsWithPreferredControlType
        }
    }
}