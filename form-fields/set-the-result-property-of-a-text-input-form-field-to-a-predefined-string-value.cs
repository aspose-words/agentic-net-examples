using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field with a name and default value.
            // Parameters: name, text type, format, default text, max length (0 = unlimited).
            builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Default", 0);

            // Save the initial document (optional, demonstrates creation step).
            doc.Save("FormFieldDemo.docx");

            // Access the form field collection and locate the field by name.
            FormField? textField = doc.Range.FormFields["MyTextField"];
            if (textField == null)
                throw new InvalidOperationException("The expected form field 'MyTextField' was not found.");

            // Set the Result property to a predefined string value.
            const string predefinedValue = "Predefined string value";
            textField.Result = predefinedValue;

            // Validate that the value was set correctly.
            if (textField.Result != predefinedValue)
                throw new InvalidOperationException("Failed to set the form field result.");

            // Save the document with the updated form field.
            doc.Save("FormFieldResult.docx");

            // Indicate successful completion.
            Console.WriteLine("Form field result set and document saved successfully.");
        }
    }
}
