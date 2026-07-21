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

        // Add some text before the target range.
        builder.Writeln("Paragraph before the target range.");

        // Insert a paragraph that will contain the checkbox.
        builder.Writeln("This paragraph will hold the checkbox form field.");

        // Retrieve the paragraph we have just added.
        // It is the last paragraph in the document body at this point.
        Paragraph targetParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph,
            doc.GetChildNodes(NodeType.Paragraph, true).Count - 1, true);

        // Move the builder cursor to the start of the target paragraph.
        builder.MoveTo(targetParagraph);

        // Insert a checkbox form field at the current position.
        // Parameters: name, default checked value, size (0 = automatic).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Set the default state of the checkbox (checked by default).
        checkBox.Default = true;

        // Optionally, set the current checked state to match the default.
        checkBox.Checked = true;

        // Add some text after the checkbox to demonstrate continuation.
        builder.Writeln();
        builder.Writeln("Paragraph after the checkbox.");

        // Save the document to the local file system.
        doc.Save("CheckboxInRange.docx");
    }
}
