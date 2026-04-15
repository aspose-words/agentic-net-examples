using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Output file name
        const string outputFile = "TextBox_RTL_Arabic.docx";

        // Create a new blank document and a builder for it
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape with a defined size
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 300, 100);

        // Move the builder's cursor inside the textbox.
        // Use the Shape's FirstParagraph (or LastParagraph) to position inside the textbox.
        builder.MoveTo(textBoxShape.FirstParagraph);

        // Set the paragraph direction to right‑to‑left (RTL)
        builder.ParagraphFormat.Bidi = true;

        // Arabic text to insert
        const string arabicText = "مرحبا بالعالم!";

        // Write the Arabic text into the textbox
        builder.Write(arabicText);

        // Save the document to disk
        doc.Save(outputFile);
    }
}
