using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Path to the markdown template file.
        string templatePath = "Template.md";

        // Read the markdown template into a string.
        string markdown = File.ReadAllText(templatePath);

        // HTML fragment that we want to insert dynamically.
        string html = "<h1 align='center'>Dynamic Title</h1><p>This paragraph is inserted from HTML.</p>";

        // Replace a placeholder in the markdown with a MERGEFIELD that will be processed as HTML.
        // The "\b Content" switch tells Aspose.Words to treat the field value as a block of content.
        string placeholder = "{{HTML}}";
        string markdownWithField = markdown.Replace(placeholder, "{{MERGEFIELD HtmlField \\b Content}}");

        // Load the modified markdown into an Aspose.Words document.
        using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(markdownWithField)))
        {
            Document doc = new Document(ms);

            // Register a callback that will insert the HTML into the merge field location.
            doc.MailMerge.FieldMergingCallback = new HtmlMergeCallback();

            // Execute mail merge, passing the HTML fragment as the field value.
            doc.MailMerge.Execute(new[] { "HtmlField" }, new object[] { html });

            // Save the final document.
            doc.Save("Result.docx");
        }
    }

    // Callback that replaces the merge field with the provided HTML.
    private class HtmlMergeCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Ensure we are handling the correct field and that it contains the "\\b" switch.
            if (args.DocumentFieldName == "HtmlField" && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Move the builder to the merge field location.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML. Use builder formatting and remove the extra empty paragraph
                // that Aspose.Words adds after block‑level HTML.
                builder.InsertHtml((string)args.FieldValue,
                    HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

                // Prevent the default text insertion.
                args.Text = string.Empty;
            }
        }

        // No image handling required for this scenario.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
