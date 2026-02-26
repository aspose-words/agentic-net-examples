using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Write("Please tick the box: ");

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name (empty string), checkedValue (false = unchecked), size (0 = auto‑size).
        FormField checkBox = builder.InsertCheckBox(string.Empty, false, 0);

        // Optional: configure additional properties of the checkbox.
        checkBox.IsCheckBoxExactSize = false;          // Use automatic sizing.
        checkBox.HelpText = "Click to toggle the box"; // Tooltip shown on F1.
        checkBox.OwnHelp = true;                       // Use the custom help text.

        // Insert a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Save the document to disk.
        doc.Save("CheckboxDocument.docx");
    }
}
