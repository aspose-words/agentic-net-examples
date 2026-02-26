using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Writeln("Please check the box below:");

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name, defaultValue, checkedValue, size (0 = automatic size).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, false, 0);

        // Set the checkbox to have an explicit size (optional).
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // size in points

        // Save the document in RTF format.
        doc.Save("CheckboxDocument.rtf", SaveFormat.Rtf);
    }
}
