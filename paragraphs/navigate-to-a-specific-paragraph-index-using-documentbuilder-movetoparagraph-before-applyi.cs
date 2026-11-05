using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample paragraphs.
        builder.Writeln("Paragraph 1: Introduction.");
        builder.Writeln("Paragraph 2: Details.");
        builder.Writeln("Paragraph 3: Conclusion.");

        // Move to the second paragraph (index 1, zero‑based) at the start of the paragraph.
        builder.MoveToParagraph(1, 0);

        // Apply formatting to the paragraph at the current position.
        // For example, make the text bold and red.
        builder.Font.Bold = true;
        builder.Font.Color = Color.Red;

        // Insert a new line to demonstrate the applied formatting.
        builder.Writeln("This line inherits the formatting of the moved paragraph.");

        // Save the document to disk.
        doc.Save("FormattedParagraph.docx");
    }
}
