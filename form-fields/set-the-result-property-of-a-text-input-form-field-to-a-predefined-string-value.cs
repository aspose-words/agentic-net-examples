using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldResultExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field with a placeholder text.
            FormField textField = builder.InsertTextInput(
                "MyTextField",                     // field name
                TextFormFieldType.Regular,         // field type
                "",                                // format (none)
                "Placeholder",                     // initial displayed text
                0);                                // no length limit

            // Validate that the field was added to the document.
            if (doc.Range.FormFields["MyTextField"] == null)
                throw new InvalidOperationException("The expected form field was not found.");

            // Set the Result property to the predefined string value.
            textField.Result = "Predefined value";

            // Save the document to disk.
            doc.Save("FormFieldResult.docx");
        }
    }
}
