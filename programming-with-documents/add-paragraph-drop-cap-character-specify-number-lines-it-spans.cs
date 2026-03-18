using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the drop cap height to span 4 lines
        builder.ParagraphFormat.LinesToDrop = 4;

        // Write the drop‑cap character (large capital letter)
        builder.Write("H");

        // Reset LinesToDrop so subsequent paragraphs behave normally
        builder.ParagraphFormat.LinesToDrop = 0;

        // Write the remaining text that will wrap around the drop cap
        builder.Writeln("ello world!");

        // Save the document to a file (adjust the path as needed)
        doc.Save("ParagraphWithDropCap.docx");
    }
}
