using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Load the DOCX template.
        Document doc = new Document("Template.docx");

        // Attach a custom field merging callback to handle check boxes and HTML.
        doc.MailMerge.FieldMergingCallback = new CustomFieldMergingCallback();

        // Define the merge fields and their values.
        string[] fieldNames = { "checkbox_Agree", "html_Description" };
        object[] fieldValues = { true, "<b>Bold description</b><br/><i>Italic text</i>" };

        // Perform a mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Render the first page of the merged document to a PNG image.
        doc.Save("Result.png", SaveFormat.Png);
    }
}

// Custom callback that inserts a check box or HTML based on the field name.
class CustomFieldMergingCallback : IFieldMergingCallback
{
    void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
    {
        // Position the builder at the current merge field.
        DocumentBuilder builder = new DocumentBuilder(args.Document);
        builder.MoveToMergeField(args.DocumentFieldName);

        if (args.DocumentFieldName.StartsWith("checkbox_"))
        {
            // Insert a check box form field. The field value is expected to be a bool.
            bool isChecked = args.FieldValue is bool b && b;
            builder.InsertCheckBox(args.DocumentFieldName, isChecked, 0);
            // Suppress the default text insertion.
            args.Text = string.Empty;
        }
        else if (args.DocumentFieldName.StartsWith("html_"))
        {
            // Insert HTML content into the document.
            string html = args.FieldValue?.ToString() ?? string.Empty;
            builder.InsertHtml(html);
            args.Text = string.Empty;
        }
        // For other fields the default behavior will apply.
    }

    void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
    {
        // No image handling required for this scenario.
    }
}
