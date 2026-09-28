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

        // Add some introductory text.
        builder.Writeln("Please fill the form below:");

        // Insert a text input form field with a placeholder (default text).
        string fieldName = "MyTextField";
        // maxLength = 0 means no length restriction.
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "Enter your name", 0);

        // Save the document with the form field.
        string initialFile = "FormFieldExample.docx";
        doc.Save(initialFile);

        // Load the document back to verify the form field.
        Document loadedDoc = new Document(initialFile);
        FormField textField = loadedDoc.Range.FormFields[fieldName];

        if (textField == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }

        // Output the current value (should be the placeholder text).
        Console.WriteLine($"Initial value of '{fieldName}': '{textField.Result}'");

        // Update the form field's value.
        string newValue = "John Doe";
        textField.Result = newValue;

        // Verify the update.
        if (textField.Result != newValue)
        {
            throw new InvalidOperationException("Failed to update the form field value.");
        }

        // Save the updated document.
        string updatedFile = "FormFieldExampleUpdated.docx";
        loadedDoc.Save(updatedFile);

        Console.WriteLine($"Form field '{fieldName}' updated to '{textField.Result}'.");
        Console.WriteLine($"Documents saved as '{initialFile}' and '{updatedFile}'.");
    }
}
