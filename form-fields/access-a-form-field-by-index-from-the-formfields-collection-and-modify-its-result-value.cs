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

        // Insert a text input form field with a default placeholder.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (none)
            "John Doe",                      // placeholder text
            50);                             // maximum length

        // Save the initial document (optional, shows the file before modification).
        doc.Save("FormFields.docx");

        // Access the form fields collection.
        FormFieldCollection fields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (fields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Retrieve the first form field by index (zero‑based).
        FormField fieldByIndex = fields[0];
        if (fieldByIndex == null)
            throw new InvalidOperationException("Form field at index 0 could not be retrieved.");

        // Modify the Result property of the text input field.
        fieldByIndex.Result = "Alice Smith";

        // Output the updated result to the console (no user interaction required).
        Console.WriteLine($"Updated field \"{fieldByIndex.Name}\" result: {fieldByIndex.Result}");

        // Save the document after modification.
        doc.Save("FormFields_Updated.docx");
    }
}
