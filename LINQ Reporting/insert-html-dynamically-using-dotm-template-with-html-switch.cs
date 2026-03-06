using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsHtmlInsertExample
{
    // Callback that inserts HTML into the document when a MERGEFIELD with the "\b" switch is encountered.
    public class HtmlMergeCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Check that the field name starts with the expected prefix and contains the "\b" switch.
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the builder to the location of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML string supplied in the data source.
                builder.InsertHtml((string)args.FieldValue);

                // We have already inserted the content, so suppress the default text insertion.
                args.Text = string.Empty;
            }
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTM template that contains a MERGEFIELD with the "\b" switch.
            const string templatePath = @"C:\Templates\HtmlInsertTemplate.dotm";

            // Load the DOTM template.
            Document doc = new Document(templatePath);

            // Ensure the template has a MERGEFIELD where HTML will be inserted.
            // If the field is not already present, insert it programmatically.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.InsertField(@"MERGEFIELD html_Content \b Content");
            builder.Writeln(); // Add a paragraph break after the field.

            // Prepare the HTML content to be merged.
            object[] mergeData = {
                "<p align='right'>Paragraph right</p>" +
                "<b>Bold text left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>"
            };

            // Assign the custom callback that will handle HTML insertion.
            doc.MailMerge.FieldMergingCallback = new HtmlMergeCallback();

            // Execute the mail merge. The field name must match the one used in the template.
            doc.MailMerge.Execute(new[] { "html_Content" }, mergeData);

            // Save the resulting document.
            const string outputPath = @"C:\Output\Result.docx";
            doc.Save(outputPath);
        }
    }
}
