using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text before the checkbox.
        builder.Writeln("Paragraph before the checkbox.");

        // Insert an empty paragraph that will hold the checkbox.
        builder.Writeln();

        // Retrieve the second paragraph (index 1) where the checkbox will be placed.
        // The first paragraph (index 0) contains the text written above.
        var targetParagraph = doc.FirstSection.Body.Paragraphs[1];

        // Move the builder's cursor to the target paragraph.
        builder.MoveTo(targetParagraph);

        // Insert a checkbox form field at the current position.
        // Parameters: name, defaultValue (initial checked state), size (0 = auto).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Set the default state of the checkbox (the value it will have when the document is opened).
        checkBox.Default = true;   // The checkbox will be checked by default.

        // Optionally, also set the current checked state to match the default.
        checkBox.Checked = true;

        // Write some text after the checkbox.
        builder.Writeln();
        builder.Writeln("Paragraph after the checkbox.");

        // Save the document to a file.
        doc.Save("CheckboxFormField.docx");
    }
}
