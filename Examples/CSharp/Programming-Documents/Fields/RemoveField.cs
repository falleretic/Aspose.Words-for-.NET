﻿using System;
using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class RemoveField : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:RemoveField
            Document doc = new Document(FieldsDir + "Field.RemoveField.doc");
            
            Field field = doc.Range.Fields[0];
            // Calling this method completely removes the field from the document
            field.Remove();
            //ExEnd:RemoveField
        }
    }
}