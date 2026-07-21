using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldReader
{
    class Program
    {
        static void Main()
        {
            // Define file name for the temporary document.
            string fileName = "FormFieldExample.docx";

            // -----------------------------------------------------------------
            // 1. Create a new document and insert a text input form field.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some introductory text.
            builder.Writeln("Please enter your name:");

            // Insert a text input form field named "MyTextField" with a placeholder.
            FormField textField = builder.InsertTextInput(
                "MyTextField",                     // field name
                TextFormFieldType.Regular,         // field type
                "",                                // default text (empty)
                "John Doe",                        // placeholder text
                50);                               // maximum length

            // Optionally set a value to demonstrate the Result property.
            textField.Result = "Alice";

            // Save the document so it can be loaded later.
            doc.Save(fileName);

            // -----------------------------------------------------------------
            // 2. Load the document from disk.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(fileName);

            // -----------------------------------------------------------------
            // 3. Retrieve the form field by its name.
            // -----------------------------------------------------------------
            FormField retrievedField = loadedDoc.Range.FormFields["MyTextField"];

            // Validate that the field exists.
            if (retrievedField == null)
                throw new InvalidOperationException("Form field 'MyTextField' was not found.");

            // -----------------------------------------------------------------
            // 4. Read and display the Result property of the field.
            // -----------------------------------------------------------------
            string fieldResult = retrievedField.Result;
            Console.WriteLine($"The result of the form field '{retrievedField.Name}' is: \"{fieldResult}\"");

            // Clean up the temporary file (optional).
            if (File.Exists(fileName))
                File.Delete(fileName);
        }
    }
}
