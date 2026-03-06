using System;
using Aspose.Words;
using Aspose.Words.Fields;

class AddCheckBoxFormField
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = "input.docx";

        // Path where the modified document will be saved.
        string outputPath = "output.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some explanatory text before the check box.
        builder.Write("Please check the box: ");

        // Insert a check box form field at the current cursor position.
        // Parameters: name, default checked value, size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Optional: set the check box to have an exact size.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Save the modified document.
        doc.Save(outputPath);
    }
}
