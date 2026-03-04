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

        // Write a prompt before the checkbox.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field.
        // Parameters: name, defaultValue, checkedValue, size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, false, 0);

        // Optionally set an explicit size for the checkbox.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Add a new paragraph after the checkbox.
        builder.InsertParagraph();

        // Save the document to a file.
        doc.Save("CheckboxFormField.docx");
    }
}
