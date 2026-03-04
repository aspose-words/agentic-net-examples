using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertCheckBoxExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some explanatory text before the checkbox.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field.
        // Parameters: name, defaultValue, checkedValue, size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, false, 0);

        // Optionally set the exact size of the checkbox.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Insert a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Save the document to a file.
        doc.Save("CheckBoxDocument.docx");
    }
}
