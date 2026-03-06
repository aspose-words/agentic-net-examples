using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains MERGEFIELDs with the \b switch.
        Document doc = new Document("Template.docx");

        // Assign a custom callback that will treat the merge data as HTML.
        doc.MailMerge.FieldMergingCallback = new HtmlMergeCallback();

        // HTML fragments to be merged into the template.
        object[] mergeData = {
            // Title field HTML.
            "<h1><span style=\"color:#0000ff; font-family:Arial;\">Hello World!</span></h1>",
            // Body field HTML.
            "<blockquote><p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p></blockquote>"
        };

        // Execute the mail merge for the fields defined in the template.
        doc.MailMerge.Execute(new[] { "html_Title", "html_Body" }, mergeData);

        // Save the populated document.
        doc.Save("Result.docx");
    }

    // Callback that inserts HTML into the document at the location of each MERGEFIELD.
    private class HtmlMergeCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Process only fields whose name starts with "html_" and that have the \b switch.
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the builder to the merge field and insert the HTML fragment.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);
                builder.InsertHtml((string)args.FieldValue);

                // Prevent the default text insertion.
                args.Text = string.Empty;
            }
        }

        // No special handling for image fields.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
