using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field with placeholder text.
        // Parameters: name, type, format, placeholder text, maxLength (0 = unlimited).
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter your name here", 0);

        // Save the document to disk.
        doc.Save("FormFieldTextInput.docx");
    }
}
