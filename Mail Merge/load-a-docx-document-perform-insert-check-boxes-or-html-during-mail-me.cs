using System;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

namespace AsposeWordsMailMergeExample
{
    // Custom callback to handle HTML insertion and checkbox insertion during mail merge.
    class HtmlAndCheckBoxMerging : IFieldMergingCallback
    {
        // Called for each merge field.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Insert HTML when the field name starts with "html_".
            if (args.DocumentFieldName.StartsWith("html_", StringComparison.OrdinalIgnoreCase))
            {
                // Move the cursor to the merge field location.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML content.
                builder.InsertHtml(args.FieldValue?.ToString() ?? string.Empty);

                // Suppress the default text insertion.
                args.Text = string.Empty;
                return;
            }

            // Insert a checkbox when the field name starts with "cb_".
            if (args.DocumentFieldName.StartsWith("cb_", StringComparison.OrdinalIgnoreCase))
            {
                // Determine the checked state (default to false if parsing fails).
                bool isChecked = false;
                if (args.FieldValue != null && bool.TryParse(args.FieldValue.ToString(), out bool parsed))
                    isChecked = parsed;

                // Move the cursor to the merge field location.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert a checkbox form field.
                // Parameters: name, isChecked, size (in points).
                builder.InsertCheckBox(args.DocumentFieldName, isChecked, 15);

                // Suppress the default text insertion.
                args.Text = string.Empty;
            }
        }

        // Not used in this scenario, but required by the interface.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX template that contains MERGEFIELDs.
            string sourceDocPath = @"C:\Docs\Template.docx";

            // Path where the resulting PNG image will be saved.
            string outputPngPath = @"C:\Docs\Result.png";

            // Load the DOCX document.
            Document doc = new Document(sourceDocPath);

            // Attach the custom field merging callback.
            doc.MailMerge.FieldMergingCallback = new HtmlAndCheckBoxMerging();

            // Define merge field names and corresponding values.
            string[] fieldNames = { "html_Title", "html_Body", "cb_Agree" };
            object[] fieldValues =
            {
                // HTML content for the title field.
                "<h1 style=\"color:#0000FF; font-family:Arial;\">Hello World!</h1>",

                // HTML content for the body field.
                "<blockquote><p>Lorem ipsum dolor sit amet, consectetur adipiscing elit.</p></blockquote>",

                // Boolean value for the checkbox field (true = checked).
                true
            };

            // Perform mail merge for a single record.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document as a PNG image (first page rendered).
            doc.Save(outputPngPath, SaveFormat.Png);
        }
    }
}
