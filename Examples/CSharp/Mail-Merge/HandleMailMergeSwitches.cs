using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class HandleMailMergeSwitches : TestDataHelper
    {
        public static void Run()
        {
            Document doc = new Document(MailMergeDir + "MailMergeSwitches.docx");

            doc.MailMerge.FieldMergingCallback = new MailMergeSwitches();

            // Fill the fields in the document with user data
            doc.MailMerge.Execute(
                new string[] { "HTML_Name" },
                new object[] { "James Bond" });

            doc.Save(ArtifactsDir + "HandleMailMergeSwitches.docx");
        }
    }

    //ExStart:HandleMailMergeSwitches
    public sealed class MailMergeSwitches : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
        {
            if (e.FieldName.StartsWith("HTML"))
            {
                if (e.Field.GetFieldCode().Contains("\\b"))
                {
                    FieldMergeField field = e.Field;

                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.DocumentFieldName, true, false);
                    builder.Write(field.TextBefore);
                    builder.InsertHtml(e.FieldValue.ToString());

                    e.Text = "";
                }
            }
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
        }
    }
    //ExEnd:HandleMailMergeSwitches
}