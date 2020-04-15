using Aspose.Words.Fields;
using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class EvaluateIFCondition
    {
        [Test]
        public static void Run()
        {
            //ExStart:EvaluateIFCondition
            DocumentBuilder builder = new DocumentBuilder();
            FieldIf field = (FieldIf) builder.InsertField("IF 1 = 1", null);

            FieldIfComparisonResult actualResult = field.EvaluateCondition();
            Console.WriteLine(actualResult);
            //ExEnd:EvaluateIFCondition
        }
    }
}