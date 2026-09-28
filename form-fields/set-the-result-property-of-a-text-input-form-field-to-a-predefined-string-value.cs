using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field with a name and a default value.
        // Parameters: name, type, format, default text, max length (0 = unlimited).
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Default Text", 0);

        // Save the initial document (optional, demonstrates creation step).
        doc.Save("FormFieldExample.docx");

        // Access the form field through the document's FormFields collection.
        FormField? formField = doc.Range.FormFields["MyTextField"];
        if (formField == null)
        {
            throw new InvalidOperationException("The expected form field 'MyTextField' was not found.");
        }

        // Set the Result property to a predefined string value.
        const string predefinedValue = "Predefined string value";
        formField.Result = predefinedValue;

        // Validate that the value was set correctly.
        if (formField.Result != predefinedValue)
        {
            throw new InvalidOperationException("Failed to set the form field result to the predefined value.");
        }

        // Save the modified document.
        doc.Save("FormFieldResult.docx");
    }
}
