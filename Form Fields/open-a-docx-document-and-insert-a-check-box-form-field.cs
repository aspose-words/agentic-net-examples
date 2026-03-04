using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        const string inputPath = "input.docx";

        // Path where the modified document will be saved.
        const string outputPath = "output.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired position).
        builder.MoveToDocumentEnd();

        // Write some explanatory text before the checkbox.
        builder.Write("Please check this box: ");

        // Insert a checkbox form field.
        // Parameters: name, checkedValue (initial state), size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Optional: enforce the exact size if you want to specify it later.
        // checkBox.IsCheckBoxExactSize = true;
        // checkBox.CheckBoxSize = 12; // size in points

        // Save the modified document.
        doc.Save(outputPath);
    }
}
