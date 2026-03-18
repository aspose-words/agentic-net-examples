using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

class RenderHtmlInParagraph
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will contain a merge field.
        // The field is marked with the "\b Content" switch so that we can replace it with HTML.
        builder.InsertField(@"MERGEFIELD  html_Content  \b Content");

        // Prepare the HTML snippets that will be merged into the document.
        // Each snippet can contain any valid HTML markup.
        object[] mergeData = {
            "<p align='center'><b>Dynamic <span style=\"color:blue;\">HTML</span> content.</b></p>",
            "<div style='border:1px solid #000; padding:5px;'>Another <i>HTML</i> block.</div>"
        };

        // Assign a custom callback that will handle the merge fields prefixed with "html_".
        // The callback inserts the HTML string directly into the document at the field location.
        doc.MailMerge.FieldMergingCallback = new HtmlMergeCallback();

        // Execute the mail merge. The field names correspond to the prefixes used above.
        doc.MailMerge.Execute(new[] { "html_Content", "html_Content" }, mergeData);

        // Save the resulting document.
        doc.Save("RenderedHtml.docx");
    }

    // Custom callback that inserts HTML into the document for fields starting with "html_".
    private class HtmlMergeCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Check that the field name starts with the expected prefix and contains the "\b" switch.
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the builder to the position of the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML string. Use the overload that respects the builder's formatting.
                builder.InsertHtml((string)args.FieldValue, useBuilderFormatting: false);

                // Since we have already inserted the content manually, suppress the default text insertion.
                args.Text = string.Empty;
            }
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this scenario.
        }
    }
}
