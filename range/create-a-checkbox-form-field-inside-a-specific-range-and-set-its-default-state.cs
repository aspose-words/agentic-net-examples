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

        // Add a paragraph that will contain the checkbox.
        builder.Writeln("Paragraph before the checkbox.");

        // Insert an empty paragraph where we will place the checkbox.
        Paragraph checkboxParagraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(checkboxParagraph);
        // Add some descriptive text before the checkbox.
        Run descriptionRun = new Run(doc, "Please check the box: ");
        checkboxParagraph.AppendChild(descriptionRun);

        // Move the builder's cursor to the start of the paragraph we just created.
        builder.MoveTo(checkboxParagraph);

        // Insert a checkbox form field at the current position.
        // Parameters: name, defaultValue, size (0 = auto size).
        FormField insertedCheckBox = builder.InsertCheckBox("MyCheckBox", false, 0);

        // Set the default state of the checkbox (checked by default).
        insertedCheckBox.Default = true;
        // Also set the current checked state to match the default.
        insertedCheckBox.Checked = true;

        // Save the document to a file.
        doc.Save("CheckboxInRange.docx");
    }
}
