using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a hanging indent of 0.25 inches (18 points) for the next paragraph.
        // Positive values create a first‑line indent, negative values create a hanging indent.
        builder.ParagraphFormat.FirstLineIndent = -18; // 0.25 inch = 18 points

        // Write the citation text. The paragraph will inherit the hanging indent set above.
        builder.Writeln("“The quick brown fox jumps over the lazy dog.” – Example Citation");

        // Ensure the output directory exists.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "HangingIndent.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputFile));

        // Save the document.
        doc.Save(outputFile);
    }
}
