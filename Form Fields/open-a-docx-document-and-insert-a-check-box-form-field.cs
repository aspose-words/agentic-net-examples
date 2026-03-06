using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document from disk.
        // Replace "input.docx" with the path to your source file.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder to edit the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name of the field, default checked status (false = unchecked), size (0 = auto).
        builder.InsertCheckBox("MyCheckBox", false, 0);

        // Save the modified document to a new file.
        // Replace "output.docx" with the desired output path.
        doc.Save("output.docx");
    }
}
