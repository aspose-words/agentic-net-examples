using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder at the very start of the main story (the document body).
        builder.MoveToDocumentStart();

        // Create a Run node that holds the text we want to insert.
        Run run = new Run(doc);
        run.Text = "Inserted text before the range.";

        // Insert the run before the current cursor position.
        builder.InsertNode(run);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
