using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        const string filePath = "FormFieldResult.docx";

        // -------------------------------------------------
        // Create a new document and add a text input field.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field with a name.
        FormField textField = builder.InsertTextInput(
            "MyTextField",                     // field name
            TextFormFieldType.Regular,         // field type
            "",                                // default text (empty)
            "Placeholder",                     // text shown when empty
            0);                                // maximum length (0 = unlimited)

        // Set the field's result so we have a value to read later.
        textField.Result = "Hello Aspose";

        // Save the document to disk.
        doc.Save(filePath);

        // -------------------------------------------------
        // Load the document and retrieve the form field.
        // -------------------------------------------------
        Document loadedDoc = new Document(filePath);

        // Access the form field collection by name.
        FormField retrievedField = loadedDoc.Range.FormFields["MyTextField"];

        // Validate that the field exists.
        if (retrievedField == null)
            throw new InvalidOperationException("Form field 'MyTextField' was not found.");

        // Read the Result property of the text input field.
        string fieldResult = retrievedField.Result ?? string.Empty;

        // Output the result to the console.
        Console.WriteLine($"Form field '{retrievedField.Name}' result: {fieldResult}");
    }
}
