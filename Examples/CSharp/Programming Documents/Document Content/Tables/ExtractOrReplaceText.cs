using System;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Tables
{
    class ExtractText : TestDataHelper
    {
        

        [Test]
        public static void ReplaceText()
        {
            //ExStart:ReplaceText
            Document doc = new Document(TablesDir + "Tables.docx");

            // Get the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Replace any instances of our string in the entire table
            table.Range.Replace("Carrots", "Eggs", new FindReplaceOptions(FindReplaceDirection.Forward));
            // Replace any instances of our string in the last cell of the table only
            table.LastRow.LastCell.Range.Replace("50", "20", new FindReplaceOptions(FindReplaceDirection.Forward));

            doc.Save(ArtifactsDir + "ReplaceText.docx");
            //ExEnd:ReplaceText
        }
    }
}