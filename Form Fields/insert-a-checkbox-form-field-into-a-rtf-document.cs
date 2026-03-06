using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some explanatory text before the checkbox.
        builder.Writeln("Please tick the box below:");

        // Insert a checkbox form field at the current cursor position.
        // Parameters: name, defaultValue (unchecked), size (0 = auto).
        builder.InsertCheckBox("MyCheckBox", false, 0);

        // Add a paragraph break after the checkbox for readability.
        builder.InsertParagraph();

        // Save the document in RTF format.
        doc.Save("CheckboxFormField.rtf");
    }
}
