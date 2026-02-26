using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.MailMerging;

namespace AsposeWordsHtmlReporting
{
    // Callback that inserts HTML into the document when a merge field with the "html_" prefix is encountered.
    class HtmlMergeCallback : IFieldMergingCallback
    {
        // Called for each merge field.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // The field must start with "html_" and contain the \b switch (block) to indicate block insertion.
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the cursor to the merge field location.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML string.
                builder.InsertHtml((string)args.FieldValue);

                // Prevent the default text insertion.
                args.Text = string.Empty;
            }
        }

        // No special handling for image fields.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a simple Word template that contains a merge field.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The merge field name starts with "html_" so the callback will process it.
            // The \b switch tells the engine that the field is a block field.
            builder.InsertField(@"MERGEFIELD html_Content \b Content");

            // Save the template (optional, just to illustrate the create‑load‑save cycle).
            template.Save("Template.docx");

            // -----------------------------------------------------------------
            // 2. Prepare the data source containing the HTML string.
            // -----------------------------------------------------------------
            var dataSource = new
            {
                // This HTML will be parsed and inserted into the document.
                html_Content = @"
                    <h1 style='color:#1E90FF;'>Hello World!</h1>
                    <p>This is a <b>bold</b> paragraph with <i>italic</i> text.</p>
                    <ul>
                        <li>Item 1</li>
                        <li>Item 2</li>
                    </ul>"
            };

            // -----------------------------------------------------------------
            // 3. Attach the field merging callback to the document.
            // -----------------------------------------------------------------
            template.MailMerge.FieldMergingCallback = new HtmlMergeCallback();

            // -----------------------------------------------------------------
            // 4. Build the report using LINQ ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The data source name ("ds") can be used inside the template if needed.
            engine.BuildReport(template, dataSource, "ds");

            // -----------------------------------------------------------------
            // 5. Save the final document.
            // -----------------------------------------------------------------
            template.Save("Result.docx");
        }
    }
}
