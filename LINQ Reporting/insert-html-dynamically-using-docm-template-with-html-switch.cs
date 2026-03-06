using System;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains MERGEFIELDs with the \b switch (HTML switch).
        Document doc = new Document("Template.docm");

        // Prepare HTML strings that will be merged into the document.
        object[] mergeData = new object[]
        {
            // Title HTML
            "<h1><span style=\"color:#0000ff; font-family:Arial;\">Hello World!</span></h1>",
            // Body HTML
            "<blockquote><p>Lorem ipsum dolor sit amet, consectetur adipiscing elit.</p></blockquote>"
        };

        // Assign a custom callback that will insert the HTML into the document.
        doc.MailMerge.FieldMergingCallback = new HtmlMergeCallback();

        // Perform the mail merge. The field names must match the MERGEFIELD names in the template.
        doc.MailMerge.Execute(new[] { "html_Title", "html_Body" }, mergeData);

        // Save the resulting document.
        doc.Save("Result.docx");
    }

    // Custom callback that handles MERGEFIELDs with the HTML switch.
    private class HtmlMergeCallback : IFieldMergingCallback
    {
        // Called for each text merge field.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Process only fields that start with "html_" and contain the \b switch.
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the builder to the location of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML string into the document.
                builder.InsertHtml((string)args.FieldValue);

                // Indicate that we have already inserted the content manually.
                args.Text = string.Empty;
            }
        }

        // No special handling for image fields.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
