using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file name for the temporary document.
        string filePath = Path.Combine(Environment.CurrentDirectory, "FormFieldExample.docx");

        // -----------------------------------------------------------------
        // 1. Create a new document and insert a text input form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some introductory text.
        builder.Writeln("Please enter your name:");

        // Insert a text input form field named "MyTextField" with a placeholder.
        FormField textField = builder.InsertTextInput(
            "MyTextField",                     // field name
            TextFormFieldType.Regular,         // field type
            "",                                // default text (empty)
            "John Doe",                        // placeholder text
            50);                               // maximum length

        // Optionally set an initial result value.
        textField.Result = "John Doe";

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
        // 4. Read the Result property of the retrieved field.
        // -----------------------------------------------------------------
        string fieldResult = retrievedField.Result;

        // Output the result to the console.
        Console.WriteLine($"The value of the form field '{retrievedField.Name}' is: {fieldResult}");

        // -----------------------------------------------------------------
        // 5. Save the document again (even though we didn't modify it,
        //    this satisfies the rule that output documents are saved).
        // -----------------------------------------------------------------
        loadedDoc.Save(filePath);
    }
}
