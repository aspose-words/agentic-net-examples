using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a hanging indent of 0.25 inches (18 points) for the next paragraph.
        // Positive FirstLineIndent = first‑line indent, negative = hanging indent.
        builder.ParagraphFormat.FirstLineIndent = -0.25 * 72; // -18 points

        // Write a sample citation paragraph.
        builder.Writeln("Doe, J. (2023). Example citation for a research paper. Journal of Examples, 12(3), 45‑67.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "HangingIndent.docx");
        doc.Save(outputPath);
    }
}
