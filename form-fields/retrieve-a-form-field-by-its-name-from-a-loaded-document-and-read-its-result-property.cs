using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldReadExample
{
    public class Program
    {
        public static void Main()
        {
            // Path for the temporary document.
            const string filePath = "FormFieldSample.docx";

            // -----------------------------------------------------------------
            // Create a new document and insert a text input form field.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Please fill in: ");

            // Insert a text input form field named "MyTextField".
            FormField textField = builder.InsertTextInput(
                "MyTextField",                     // field name
                TextFormFieldType.Regular,         // field type
                "",                                // default text (empty)
                "Default placeholder",             // placeholder text
                0);                                // max length (0 = unlimited)

            // Set an initial value for demonstration purposes.
            textField.Result = "Sample value";

            // Save the document so it can be loaded later.
            doc.Save(filePath);

            // -----------------------------------------------------------------
            // Load the document from disk.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(filePath);

            // -----------------------------------------------------------------
            // Retrieve the form field by its name and read its Result.
            // -----------------------------------------------------------------
            FormField retrievedField = loadedDoc.Range.FormFields["MyTextField"];
            if (retrievedField == null)
                throw new InvalidOperationException("Form field 'MyTextField' was not found in the document.");

            // The Result property may be null; guard against it.
            string fieldResult = retrievedField.Result ?? string.Empty;

            // Output the result to the console.
            Console.WriteLine($"Result of form field '{retrievedField.Name}': {fieldResult}");
        }
    }
}
