using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field with a specific name.
        const string fieldName = "MyTextField";
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(
            fieldName,
            TextFormFieldType.Regular,
            "",
            "John Doe",   // default placeholder text
            50);

        // Set an explicit value for demonstration purposes.
        textField.Result = "Alice";

        // Save the document locally.
        const string filePath = "FormFieldSample.docx";
        doc.Save(filePath);

        // Load the document back from disk.
        Document loadedDoc = new Document(filePath);

        // Retrieve the form field by its name.
        FormField retrievedField = loadedDoc.Range.FormFields[fieldName];
        if (retrievedField == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }

        // Read the Result property of the form field.
        string fieldResult = retrievedField.Result;

        // Output the result to the console.
        Console.WriteLine($"Form field '{fieldName}' result: {fieldResult}");
    }
}
