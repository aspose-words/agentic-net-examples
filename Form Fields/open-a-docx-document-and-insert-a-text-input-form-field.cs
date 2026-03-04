using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\Template.docx";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\Template_WithFormField.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: add some explanatory text before the form field.
        builder.Write("Please enter your name: ");

        // Insert a text input form field at the current cursor position.
        // Parameters:
        //   name        – bookmark/form field name (empty string creates no bookmark).
        //   type        – type of the text form field (regular text in this case).
        //   format      – format string (not needed for regular text, so empty).
        //   fieldValue  – placeholder text shown to the user.
        //   maxLength   – 0 means unlimited length.
        builder.InsertTextInput(
            name: "MyTextInput",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Enter your name here",
            maxLength: 0);

        // Save the modified document.
        doc.Save(outputPath);
    }
}
