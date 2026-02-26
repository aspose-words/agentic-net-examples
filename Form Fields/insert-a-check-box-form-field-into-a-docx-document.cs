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
        builder.Write("Please tick the box if you agree: ");

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name, checkedValue (initial state), size (0 = auto size).
        FormField checkBox = builder.InsertCheckBox("AgreementCheckBox", false, 0);

        // Optionally set additional properties.
        checkBox.IsCheckBoxExactSize = true;   // Use exact size if needed.
        checkBox.CheckBoxSize = 12;            // Size in points (effective because IsCheckBoxExactSize is true).
        checkBox.HelpText = "Click to agree";
        checkBox.OwnHelp = true;

        // Insert a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Save the document to a DOCX file.
        doc.Save("CheckboxFormField.docx");
    }
}
