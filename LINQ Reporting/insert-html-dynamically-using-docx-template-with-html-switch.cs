using System;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

namespace HtmlDynamicInsertExample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Create an instance of the helper class and run the example.
            var example = new HtmlDynamicInsert();
            example.Execute();
        }
    }

    public class HtmlDynamicInsert
    {
        public void Execute()
        {
            // Load the DOCX template that contains MERGEFIELDs with the \b switch.
            Document doc = new Document("Template.docx");

            // Register a callback that will replace the merge fields with HTML content.
            doc.MailMerge.FieldMergingCallback = new HtmlFieldMergingCallback();

            // HTML fragments to be merged into the template.
            object[] mergeData = new object[]
            {
                "<h2 style='color:blue;'>Welcome</h2><p>This is <b>bold</b> text.</p>",
                "<ul><li>Item 1</li><li>Item 2</li></ul>"
            };

            // Execute mail merge. Field names must match those defined in the template
            // (e.g., html_Title and html_Body).
            doc.MailMerge.Execute(new[] { "html_Title", "html_Body" }, mergeData);

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }

    // Callback that inserts HTML into MERGEFIELDs whose name starts with "html_"
    // and contain the \b switch (the "HTML switch").
    internal class HtmlFieldMergingCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
            {
                // Position the builder at the merge field.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the HTML fragment.
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
