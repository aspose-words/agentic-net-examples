using System;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Drawing;

namespace AsposeWordsReporting
{
    // Callback that inserts HTML into the document when a merge field with the "html_" prefix is encountered.
    class HtmlFieldMergingCallback : IFieldMergingCallback
    {
        // Called for each merge field during the mail‑merge operation.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // We only handle fields whose name starts with "html_".
            if (args.DocumentFieldName.StartsWith("html_", StringComparison.OrdinalIgnoreCase))
            {
                // Move the builder to the location of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML string (the field value) at that position.
                // The field value is expected to be a string containing valid HTML.
                builder.InsertHtml(args.FieldValue?.ToString() ?? string.Empty);

                // Clear the default text that Aspose.Words would otherwise insert.
                args.Text = string.Empty;
            }
        }

        // No special handling for image fields in this scenario.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTM template that contains merge fields like <<[html_Title]>> and <<[html_Body]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.dotm";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source. The property names must match the merge field names.
            var dataSource = new
            {
                html_Title = "<h1 style=\"color:#2E86C1;\">Quarterly Report</h1>",
                html_Body  = "<p>This quarter we achieved a <b>15% growth</b> in revenue.</p>" +
                             "<ul><li>Product A: +20%</li><li>Product B: +10%</li></ul>"
            };

            // Assign the custom callback that will turn the merge field values into HTML.
            doc.MailMerge.FieldMergingCallback = new HtmlFieldMergingCallback();

            // Execute the mail‑merge. The field names array must be in the same order as the values array.
            doc.MailMerge.Execute(
                new[] { "html_Title", "html_Body" },
                new object[] { dataSource.html_Title, dataSource.html_Body });

            // Save the resulting document.
            const string outputPath = @"C:\Reports\QuarterlyReport.docx";
            doc.Save(outputPath);
        }
    }
}
