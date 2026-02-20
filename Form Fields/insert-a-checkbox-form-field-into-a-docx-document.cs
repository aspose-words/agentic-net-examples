using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertCheckBoxFormField
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and the checkbox form field.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some explanatory text before the checkbox.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field.
        // Parameters: name, default checked state, size (points).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);

        // Optional: make the size exact, set help/status text, etc.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.HelpText = "Right‑click to toggle";
        checkBox.OwnHelp = true;
        checkBox.StatusText = "Checkbox status";
        checkBox.OwnStatus = true;

        // Save the document to a DOCX file.
        doc.Save("CheckboxFormField.docx");
    }
}
