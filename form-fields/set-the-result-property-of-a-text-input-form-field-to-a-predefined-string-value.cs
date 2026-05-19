using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field with a placeholder.
        // Name: "MyTextInput", Type: Regular, no format, placeholder text, unlimited length (0).
        FormField textField = builder.InsertTextInput(
            "MyTextInput",
            TextFormFieldType.Regular,
            "",
            "Placeholder text",
            0);

        // Ensure the form field was created.
        if (textField == null)
            throw new InvalidOperationException("Failed to create the text input form field.");

        // Set the Result property to a predefined string value.
        textField.Result = "Predefined string value";

        // Save the document to disk.
        doc.Save("ResultFormField.docx");
    }
}
