using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field with placeholder text.
        // Name: "TextInput", type: Regular, no format, placeholder: "Placeholder text", unlimited length (0).
        FormField textField = builder.InsertTextInput(
            "TextInput",
            TextFormFieldType.Regular,
            "",
            "Placeholder text",
            0);

        // Validate that the form field exists.
        if (doc.Range.FormFields["TextInput"] == null)
            throw new InvalidOperationException("The text input form field was not created.");

        // Save the document to disk.
        doc.Save("FormFieldExample.docx");
    }
}
