﻿using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Fields
{
    class InsertAdvanceFieldWithOutDocumentBuilder : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:InsertAdvanceFieldWithOutDocumentBuilder
            Document doc = new Document(FieldsDir + "in.doc");
            // Get paragraph you want to append this Advance field to
            Paragraph para = (Paragraph) doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert an Advance field like this:
            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }

            // Create instance of FieldAdvance class and lets build the above field code
            FieldAdvance field = (FieldAdvance) para.AppendField(FieldType.FieldAdvance, false);
            
            // { ADVANCE \\d 10 " }
            field.DownOffset = "10";

            // { ADVANCE \\d 10 \\l 10 }
            field.LeftOffset = "10";

            // { ADVANCE \\d 10 \\l 10 \\r -3.3 }
            field.RightOffset = "-3.3";

            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
            field.UpOffset = "0";

            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
            field.HorizontalPosition = "100";

            // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
            field.VerticalPosition = "100";

            // Finally update this Advance field
            field.Update();

            doc.Save(ArtifactsDir + "InsertAdvanceFieldWithOutDocumentBuilder.doc");
            //ExEnd:InsertAdvanceFieldWithOutDocumentBuilder
        }
    }
}