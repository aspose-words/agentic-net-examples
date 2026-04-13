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

        // Insert a text input form field with a name and placeholder text.
        builder.Write("Please enter a value: ");
        FormField textField = builder.InsertTextInput(
            "MyTextField",                     // field name
            TextFormFieldType.Regular,         // field type
            "",                                // default text (empty)
            "Placeholder",                     // placeholder text shown when empty
            50);                               // maximum length

        // Optional: set a default value and a text format.
        textField.TextInputDefault = "Default value";
        textField.TextInputFormat = "UPPERCASE";

        // Access the form fields collection.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Retrieve the first form field by index (0‑based).
        FormField fieldByIndex = formFields[0];
        if (fieldByIndex == null)
            throw new InvalidOperationException("Form field at index 0 could not be retrieved.");

        // Modify the Result property of the text form field.
        const string newResult = "Updated result";
        fieldByIndex.Result = newResult;

        // Verify that the value was updated correctly.
        if (fieldByIndex.Result != newResult)
            throw new InvalidOperationException("Failed to update the form field result.");

        // Save the document to disk.
        doc.Save("FormFieldResult.docx");

        // Indicate successful completion.
        Console.WriteLine("Form field updated and document saved.");
    }
}
