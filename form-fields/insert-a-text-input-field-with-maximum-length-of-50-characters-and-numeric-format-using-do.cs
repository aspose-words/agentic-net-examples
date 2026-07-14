using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a label before the form field.
            builder.Write("Enter a number (max 50 characters): ");

            // Insert a numeric text input form field.
            // Parameters: name, type, format string, default value, maximum length.
            FormField numberField = builder.InsertTextInput(
                "NumberField",
                TextFormFieldType.Number,
                "",      // No custom format string.
                "0",     // Default placeholder value.
                50);     // Maximum length of the field.

            // Optional: set help and status text for the field.
            numberField.HelpText = "Only numeric input is allowed.";
            numberField.StatusText = "Enter a numeric value.";

            // Save the document to a file.
            doc.Save("FormField.docx");
        }
    }
}
