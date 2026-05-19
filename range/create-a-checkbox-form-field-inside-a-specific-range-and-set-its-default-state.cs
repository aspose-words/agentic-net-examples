using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Document with a checkbox form field inserted into a specific range.");

        // Insert a paragraph that will serve as the target range.
        builder.Writeln("Target paragraph: ");

        // Retrieve the last paragraph (the one we just added).
        Paragraph targetParagraph = doc.LastSection.Body.Paragraphs[doc.LastSection.Body.Paragraphs.Count - 1];

        // Move the builder's cursor to the start of the target paragraph.
        builder.MoveTo(targetParagraph);

        // Insert a checkbox form field at the current position.
        // Parameters: name, defaultValue (checked state), size (0 = auto).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Set the default value of the checkbox (the state it will have when the document is opened).
        checkBox.Default = true; // The checkbox will be checked by default.

        // Optionally, set the current checked state (can differ from the default).
        checkBox.Checked = false; // Currently unchecked.

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CheckboxInRange.docx");
        doc.Save(outputPath);
    }
}
