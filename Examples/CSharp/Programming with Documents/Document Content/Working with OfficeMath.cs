using Aspose.Words.Math;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Document_Content
{
    class WorkingWithOfficeMath : TestDataHelper
    {
        [Test]
        public static void SpecifyLocaleAtFieldlevel()
        {
            //ExStart:SpecifylocaleAtFieldlevel
            Document doc = new Document(MyDir + "Office math.docx");
            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);

            // Gets/sets Office Math display format type which represents whether an equation is displayed inline with the text or displayed on its own line
            officeMath.DisplayType = OfficeMathDisplayType.Display; // or OfficeMathDisplayType.Inline

            // Gets/sets Office Math justification
            officeMath.Justification = OfficeMathJustification.Left; // Left justification of Math Paragraph

            doc.Save(ArtifactsDir + "MathEquations.docx");
            //ExEnd:SpecifylocaleAtFieldlevel
        }
    }
}