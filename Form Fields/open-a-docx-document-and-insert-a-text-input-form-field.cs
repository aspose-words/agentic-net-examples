using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertTextInputFormField
{
    static void Main()
    {
        // Path to the existing DOCX document.
        string inputPath = "input.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field at the current cursor position.
        // Parameters: name, type, format, default text, max length (0 = unlimited).
        builder.InsertTextInput(
            name: "MyTextInput",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Enter your text here",
            maxLength: 0);

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
