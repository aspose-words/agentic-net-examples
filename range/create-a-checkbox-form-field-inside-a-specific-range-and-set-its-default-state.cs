using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Paragraph before the checkbox.");
        builder.Writeln("Insert checkbox in this paragraph:");

        // Retrieve the second paragraph (index 1) where we want to place the checkbox.
        Paragraph targetParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);

        // Move the builder's cursor to the start of the target paragraph.
        builder.MoveTo(targetParagraph);

        // Insert a checkbox form field at the current position.
        // Parameters: name, defaultValue, checkedValue, size (0 = auto).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, true, 0);

        // Set explicit size if desired.
        checkBox.IsCheckBoxExactSize = true;
        checkBox.CheckBoxSize = 12; // points

        // Save the document.
        doc.Save("CheckboxInRange.docx");
    }
}
