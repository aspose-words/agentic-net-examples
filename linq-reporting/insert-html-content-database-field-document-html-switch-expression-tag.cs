using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

namespace AsposeWordsHtmlInsertExample
{
    // Custom callback that inserts HTML content for merge fields whose name starts with "html_"
    // and that contain the "\b" switch (block formatting switch).
    class HtmlMergeCallback : IFieldMergingCallback
    {
        // Called for each simple merge field.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Check for the "html_" prefix and the presence of the block switch.
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the builder to the location of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML string from the data source.
                // The field value is expected to be a string containing HTML markup.
                builder.InsertHtml((string)args.FieldValue);

                // We have already inserted the content manually, so suppress the default text insertion.
                args.Text = string.Empty;
            }
        }

        // Called for image merge fields – not used in this example.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No action required.
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Insert two MERGEFIELDs that will receive HTML content.
            // The "\b Content" switch tells Aspose.Words to treat the field as a block.
            builder.InsertField(@"MERGEFIELD  html_Title  \b Content");
            builder.InsertParagraph(); // Add a paragraph break between fields.
            builder.InsertField(@"MERGEFIELD  html_Body  \b Content");

            // 3. Prepare sample HTML data that would normally come from a database.
            object[] htmlData = new object[]
            {
                // Title HTML
                "<h1 style=\"color:#1E90FF; font-family:Arial;\">Welcome to Aspose.Words</h1>",

                // Body HTML
                "<p>This example demonstrates how to <b>insert HTML</b> from a data source " +
                "into a Word document using the <i>html</i> switch in a MERGEFIELD.</p>" +
                "<ul><li>Item 1</li><li>Item 2</li></ul>"
            };

            // 4. Assign the custom callback to handle HTML insertion.
            doc.MailMerge.FieldMergingCallback = new HtmlMergeCallback();

            // 5. Execute the mail merge. Field names correspond to the MERGEFIELDs inserted above.
            doc.MailMerge.Execute(new[] { "html_Title", "html_Body" }, htmlData);

            // 6. Save the resulting document.
            doc.Save("HtmlInsertResult.docx");
        }
    }
}
