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

            // Insert a text input form field with a name.
            builder.Write("Please enter your name: ");
            FormField textField = builder.InsertTextInput(
                name: "NameField",
                type: TextFormFieldType.Regular,
                format: "",
                fieldValue: "",
                maxLength: 0);

            // Ensure the form field was created.
            if (textField == null)
                throw new InvalidOperationException("Failed to create the text input form field.");

            // Set the Result property to a predefined string.
            const string predefinedValue = "John Doe";
            textField.Result = predefinedValue;

            // Validate that the value was set correctly.
            if (textField.Result != predefinedValue)
                throw new InvalidOperationException("The form field result was not set correctly.");

            // Save the document to disk.
            const string outputPath = "FormFieldsResult.docx";
            doc.Save(outputPath);
        }
    }
}
