using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.DocumentEx
{
    class UpdateContentControls : TestDataHelper
    {
        [Test]
        public static void SetCurrentStateOfCheckBox()
        {
            //ExStart:SetCurrentStateOfCheckBox
            // Open an existing document
            Document doc = new Document(DocumentDir + "Structured document tags.docx");
            
            // Get the first content control from the document
            Markup.StructuredDocumentTag sdtCheckBox =
                (Markup.StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            // StructuredDocumentTag.Checked property gets/sets current state of the Checkbox SDT
            if (sdtCheckBox.SdtType == SdtType.Checkbox)
                sdtCheckBox.Checked = true;

            doc.Save(ArtifactsDir + "SetCurrentStateOfCheckBox.docx");
            //ExEnd:SetCurrentStateOfCheckBox
        }

        [Test]
        public static void ModifyContentControls()
        {
            //ExStart:ModifyContentControls
            // Open an existing document
            Document doc = new Document(DocumentDir + "Structured document tags.docx");

            foreach (Markup.StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                switch (sdt.SdtType)
                {
                    case SdtType.PlainText:
                    {
                        sdt.RemoveAllChildren();
                        Paragraph para = sdt.AppendChild(new Paragraph(doc)) as Paragraph;
                        Run run = new Run(doc, "new text goes here");
                        para.AppendChild(run);
                        break;
                    }
                    case SdtType.DropDownList:
                    {
                        SdtListItem secondItem = sdt.ListItems[2];
                        sdt.ListItems.SelectedValue = secondItem;
                        break;
                    }
                    case SdtType.Picture:
                    {
                        Shape shape = (Shape) sdt.GetChild(NodeType.Shape, 0, true);
                        if (shape.HasImage)
                        {
                            shape.ImageData.SetImage(DocumentDir + "Watermark.png");
                        }

                        break;
                    }
                }
            }
            
            doc.Save(ArtifactsDir + "ModifyContentControls.docx");
            //ExEnd:ModifyContentControls
        }
    }
}