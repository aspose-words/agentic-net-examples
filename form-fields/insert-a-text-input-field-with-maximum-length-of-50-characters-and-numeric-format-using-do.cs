using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a numeric text input form field.
            // Parameters: name, type, format, default value, maximum length.
            // TextFormFieldType.Number restricts input to numbers.
            // Format "0" will display the number without extra formatting.
            FormField numericField = builder.InsertTextInput(
                "NumericInput",
                TextFormFieldType.Number,
                "0",
                "0",
                50);

            // Verify that the field was added and its properties are correct.
            FormField? retrievedField = doc.Range.FormFields["NumericInput"];
            if (retrievedField == null)
                throw new InvalidOperationException("The form field was not found.");

            if (retrievedField.MaxLength != 50)
                throw new InvalidOperationException("MaxLength is not set correctly.");

            // Set a sample numeric value using the appropriate method.
            retrievedField.SetTextInputValue(12345);

            // Save the document to disk.
            doc.Save("NumericInputForm.docx");
        }
    }
}
