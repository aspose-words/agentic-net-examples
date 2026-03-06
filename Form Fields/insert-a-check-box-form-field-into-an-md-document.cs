using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the checkbox.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name (empty), checkedValue (false), size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox(string.Empty, false, 0);

        // Optionally set an explicit size for the checkbox.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Insert a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Save the document to a file.
        doc.Save("CheckBoxDocument.docx");
    }
}
