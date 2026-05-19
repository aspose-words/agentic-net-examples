using System;
using Aspose.Words;
using Aspose.Words.Fields;   // Required for FormField and TextFormFieldType

public class Program
{
    public static void Main()
    {
        // Path for the sample document.
        const string filePath = "FormFieldSample.docx";

        // -----------------------------------------------------------------
        // 1. Create a new document and insert a text input form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some introductory text.
        builder.Writeln("Please enter your name:");

        // Insert a text input form field named "MyTextField".
        // Parameters: name, type, default text, placeholder text, max length.
        FormField textField = builder.InsertTextInput(
            "MyTextField",
            TextFormFieldType.Regular,
            "",
            "John Doe",
            50);

        // Save the document so it can be loaded later.
        doc.Save(filePath);

        // -----------------------------------------------------------------
        // 2. Load the document from disk.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(filePath);

        // -----------------------------------------------------------------
        // 3. Retrieve the form field by its name.
        // -----------------------------------------------------------------
        FormField retrievedField = loadedDoc.Range.FormFields["MyTextField"];

        // Validate that the field exists.
        if (retrievedField == null)
        {
            throw new InvalidOperationException("Form field 'MyTextField' was not found in the document.");
        }

        // -----------------------------------------------------------------
        // 4. Read the Result property of the form field.
        // -----------------------------------------------------------------
        string fieldResult = retrievedField.Result;

        // Output the result to the console.
        Console.WriteLine($"Result of form field '{retrievedField.Name}': \"{fieldResult}\"");

        // Save the processed document.
        loadedDoc.Save("FormFieldSample_Processed.docx");
    }
}
