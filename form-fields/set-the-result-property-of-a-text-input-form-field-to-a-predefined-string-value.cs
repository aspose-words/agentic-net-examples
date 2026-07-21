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

        // Insert a text input form field with a placeholder.
        FormField textField = builder.InsertTextInput(
            name: "MyTextInput",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Placeholder",
            maxLength: 0);

        // Verify that the form field was created.
        if (textField == null)
            throw new InvalidOperationException("Failed to create the text input form field.");

        // Set the Result property to the predefined string.
        const string predefinedValue = "Hello Aspose!";
        textField.Result = predefinedValue;

        // Validate that the Result was set correctly.
        if (textField.Result != predefinedValue)
            throw new InvalidOperationException("The Result property was not set as expected.");

        // Save the document to disk.
        doc.Save("FormFieldResult.docx");
    }
}
