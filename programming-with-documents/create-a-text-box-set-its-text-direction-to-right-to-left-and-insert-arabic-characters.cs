using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape with a specific size.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Move the builder cursor inside the text box.
        builder.MoveTo(textBox.FirstParagraph);

        // Set the paragraph direction to right‑to‑left (Bidi) for Arabic text.
        builder.ParagraphFormat.Bidi = true;

        // Insert Arabic characters into the text box.
        builder.Writeln("مرحبا بالعالم");

        // Reset the Bidi setting for subsequent content (optional).
        builder.ParagraphFormat.Bidi = false;

        // Save the document to disk.
        string outputPath = "TextBox_RTL.docx";
        doc.Save(outputPath);

        // Verify that the file was created and can be reopened.
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // No further actions needed; the program will exit.
        }
    }
}
