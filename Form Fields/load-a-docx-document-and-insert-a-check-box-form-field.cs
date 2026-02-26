using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document from the file system.
        Document doc = new Document("input.docx");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name of the field, default checked status, size in points (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Optional: set the exact size of the checkbox and enable exact sizing.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Save the modified document to a new file.
        doc.Save("output.docx");
    }
}
