using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some explanatory text before the checkbox
        builder.Write("Please tick the box if you agree: ");

        // Insert a checkbox form field at the current cursor position
        // Parameters: name, isChecked, size (0 = auto)
        FormField checkBox = builder.InsertCheckBox("AgreeCheckBox", false, 0);

        // Optionally set additional properties on the checkbox
        checkBox.IsCheckBoxExactSize = true;   // Use exact size if needed
        checkBox.CheckBoxSize = 12;            // Size in points (effective when IsCheckBoxExactSize is true)

        // Insert a paragraph break after the checkbox
        builder.InsertParagraph();

        // Save the document to a file
        doc.Save("CheckboxFormField.docx");
    }
}
