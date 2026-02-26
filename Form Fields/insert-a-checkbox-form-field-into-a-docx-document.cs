using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some explanatory text before the checkbox.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field.
        // Parameters: name, checkedValue (false = unchecked), size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Optional: make the size exact if you want to specify it.
        // checkBox.IsCheckBoxExactSize = true;
        // checkBox.CheckBoxSize = 20; // size in points

        // Insert a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Save the document to a DOCX file.
        doc.Save("CheckboxFormField.docx");
    }
}
