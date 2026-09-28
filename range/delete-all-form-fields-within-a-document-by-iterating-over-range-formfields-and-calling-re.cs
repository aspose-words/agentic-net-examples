using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample form fields.
        // Use TextFormFieldType enum for the text input field type.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Sample text", 0);
        builder.Writeln();
        builder.InsertCheckBox("CheckBox", true, 0);
        builder.Writeln();

        // Save the original document (optional, demonstrates creation).
        doc.Save("Original.docx");

        // Delete all form fields by iterating over the collection and removing each.
        FormFieldCollection formFields = doc.Range.FormFields;
        while (formFields.Count > 0)
        {
            // Remove the first form field; the collection updates automatically.
            formFields[0].Remove();
        }

        // Save the document after removal of form fields.
        doc.Save("Cleaned.docx");
    }
}
