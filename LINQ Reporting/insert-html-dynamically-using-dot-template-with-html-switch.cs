using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

class InsertHtmlWithSwitch
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD whose name starts with "html_" and includes the \b switch.
        // The \b switch tells the callback that the field value should be treated as HTML.
        builder.InsertField(@"MERGEFIELD html_Body \b Content");

        // Prepare HTML strings that will be merged into the document.
        object[] mergeData = new object[]
        {
            "<h1 style=\"color:blue;\">Dynamic Title</h1>",          // html_Title
            "<p>This paragraph is <b>bold</b> and <i>italic</i>.</p>" // html_Body
        };

        // Register a custom callback that will insert the HTML into the document.
        doc.MailMerge.FieldMergingCallback = new HtmlFieldMergingCallback();

        // Execute the mail merge. Field names must match the MERGEFIELD names.
        doc.MailMerge.Execute(new[] { "html_Title", "html_Body" }, mergeData);

        // Save the resulting document.
        doc.Save("InsertHtmlWithSwitch.docx");
    }

    // Callback that handles fields with the "html_" prefix and the \b switch.
    private class HtmlFieldMergingCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Check for the prefix and the presence of the \b switch.
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the builder to the location of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML string. Use the overload that takes only the HTML text.
                builder.InsertHtml((string)args.FieldValue);

                // Suppress the default text insertion.
                args.Text = string.Empty;
            }
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this scenario.
        }
    }
}
