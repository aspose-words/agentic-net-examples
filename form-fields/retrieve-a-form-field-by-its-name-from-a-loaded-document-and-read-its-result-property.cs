using System;
using Aspose.Words;
using Aspose.Words.Fields; // Provides FormField, TextFormFieldType, etc.

namespace FormFieldReader
{
    public class Program
    {
        public static void Main()
        {
            // Path for the temporary document.
            const string filePath = "FormFields.docx";

            // 1. Create a new document and insert a text input form field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field named "MyTextField" with default value "Default".
            // Use the overload that requires TextFormFieldType and maxLength.
            builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Default", 0);

            // Save the document so it can be loaded later.
            doc.Save(filePath);

            // 2. Load the document from disk.
            Document loadedDoc = new Document(filePath);

            // 3. Retrieve the form field by its name.
            FormField formField = loadedDoc.Range.FormFields["MyTextField"];
            if (formField == null)
            {
                throw new InvalidOperationException("Form field 'MyTextField' was not found in the document.");
            }

            // 4. Read the Result property (the current value of the text input field).
            string fieldResult = formField.Result;

            // Output the result to the console.
            Console.WriteLine($"The result of the form field 'MyTextField' is: \"{fieldResult}\"");

            // 5. Save the document after reading (optional but satisfies saving rule).
            loadedDoc.Save("FormFieldsResult.docx");
        }
    }
}
