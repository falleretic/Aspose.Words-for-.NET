using Aspose.Words.Replacing;
using System;
using System.Text.RegularExpressions;
using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class UsingLegacyOrder : TestDataHelper
    {
        public static void Run()
        {
            FineReplaceUsingLegacyOrder();
        }

        //ExStart:FineReplaceUsingLegacyOrder
        public static void FineReplaceUsingLegacyOrder()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert 3 tags to appear in sequential order, the second of which will be inside a text box
            builder.Writeln("[tag 1]");
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
            builder.Writeln("[tag 3]");

            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("[tag 2]");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplacingCallback();
            options.UseLegacyOrder = true;

            doc.Range.Replace(new Regex(@"\[(.*?)\]"), "", options);

            doc.Save(ArtifactsDir + "FineReplaceUsingLegacyOrder.docx");
        }

        private class ReplacingCallback : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Console.Write(e.Match.Value);
                return ReplaceAction.Replace;
            }
        }
        //ExEnd:FineReplaceUsingLegacyOrder
    }
}