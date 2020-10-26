using Aspose.Words.Fields;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class MailMergeFormFields : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:MailMergeFormFields
            Document doc = new Document(MyDir + "Mail merge destinations - Fax.docx");

            // Setup mail merge event handler to do the custom work
            doc.MailMerge.FieldMergingCallback = new HandleMergeField();

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // This is the data for mail merge
            string[] fieldNames = {
                "RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
                "Subject", "Body", "Urgent", "ForReview", "PleaseComment"
            };

            object[] fieldValues = {
                "Josh", "Jenny", "123456789", "", "Hello",
                "<b>HTML Body Test message 1</b>", true, false, true
            };

            // Execute the mail merge
            doc.MailMerge.Execute(fieldNames, fieldValues);

            doc.Save(ArtifactsDir + "MailMergeFormFields.docx");
            //ExEnd:MailMergeFormFields
        }

        //ExStart:HandleMergeField
        private class HandleMergeField : IFieldMergingCallback
        {
            /// <summary>
            /// This handler is called for every mail merge field found in the document,
            ///  for every record found in the data source.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (mBuilder == null)
                    mBuilder = new DocumentBuilder(e.Document);

                // We decided that we want all boolean values to be output as check box form fields
                if (e.FieldValue is bool)
                {
                    // Move the "cursor" to the current merge field
                    mBuilder.MoveToMergeField(e.FieldName);

                    // It is nice to give names to check boxes. Lets generate a name such as MyField21 or so
                    string checkBoxName = $"{e.FieldName}{e.RecordIndex}";

                    // Insert a check box
                    mBuilder.InsertCheckBox(checkBoxName, (bool) e.FieldValue, 0);

                    // Nothing else to do for this field
                    return;
                }

                switch (e.FieldName)
                {
                    // We want to insert html during mail merge
                    case "Body":
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.InsertHtml((string) e.FieldValue);
                        break;
                    // Another example, we want the Subject field to come out as text input form field
                    case "Subject":
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        string textInputName = $"{e.FieldName}{e.RecordIndex}";
                        mBuilder.InsertTextInput(textInputName, TextFormFieldType.Regular, "", (string) e.FieldValue, 0);
                        break;
                    }
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing
            }

            private DocumentBuilder mBuilder;
        }
        //ExEnd:HandleMergeField
    }
}