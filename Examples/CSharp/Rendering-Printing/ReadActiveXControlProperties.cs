using System;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class ReadActiveXControlProperties : TestDataHelper
    {
        public static void Run()
        {
            Document doc = new Document(MailMergeDir + "ActiveXControl.docx");

            string properties = "";
            // Retrieve shapes from the document
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                OleControl oleControl = shape.OleFormat.OleControl;
                if (oleControl.IsForms2OleControl)
                {
                    Forms2OleControl checkBox = (Forms2OleControl) oleControl;
                    properties = properties + "\nCaption: " + checkBox.Caption;
                    properties = properties + "\nValue: " + checkBox.Value;
                    properties = properties + "\nEnabled: " + checkBox.Enabled;
                    properties = properties + "\nType: " + checkBox.Type;
                    if (checkBox.ChildNodes != null)
                    {
                        properties = properties + "\nChildNodes: " + checkBox.ChildNodes;
                    }

                    properties += "\n";
                }
            }

            properties = properties + "\nTotal ActiveX Controls found: " + doc.GetChildNodes(NodeType.Shape, true).Count;
            Console.WriteLine("\n" + properties);
        }
    }
}