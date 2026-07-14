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
        // Name: "MyTextInput", type: regular text, no format, placeholder text, unlimited length (0).
        FormField textField = builder.InsertTextInput(
            "MyTextInput",
            TextFormFieldType.Regular,
            "",
            "Placeholder text",
            0);

        // Ensure the form field was created.
        if (textField == null)
            throw new InvalidOperationException("Failed to create the text input form field.");

        // Access the form field via the collection to demonstrate reading by name.
        FormField retrievedField = doc.Range.FormFields["MyTextInput"];
        if (retrievedField == null)
            throw new InvalidOperationException("The expected form field 'MyTextInput' was not found.");

        // Set the Result property to a predefined string value.
        const string predefinedValue = "Hello Aspose.Words!";
        retrievedField.Result = predefinedValue;

        // Verify that the value was set correctly.
        if (retrievedField.Result != predefinedValue)
            throw new InvalidOperationException("Failed to set the Result property of the form field.");

        // Save the document to disk.
        doc.Save("ResultFormField.docx");
    }
}
