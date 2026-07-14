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

        // Insert a text input form field with a placeholder.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",                 // field name
            TextFormFieldType.Regular,  // field type
            "",                          // default text (none)
            "",                          // format (none)
            50);                         // maximum length

        // Ensure that the document contains at least one form field.
        FormFieldCollection fields = doc.Range.FormFields;
        if (fields == null || fields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Access the first form field by its zero‑based index.
        FormField fieldByIndex = fields[0];
        if (fieldByIndex == null)
            throw new InvalidOperationException("Form field at index 0 could not be retrieved.");

        // Modify the Result property of the form field.
        const string newValue = "Jane Smith";
        fieldByIndex.Result = newValue;

        // Verify that the value was updated correctly.
        if (fieldByIndex.Result != newValue)
            throw new InvalidOperationException("Failed to update the form field result.");

        // Save the modified document.
        doc.Save("FormFieldModified.docx");
    }
}
