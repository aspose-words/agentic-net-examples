using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a hanging indent of 0.25 inches (18 points). Negative value creates a hanging indent.
        builder.ParagraphFormat.FirstLineIndent = -18; // 0.25 inch = 18 points

        // Add the citation paragraph.
        builder.Writeln("Doe, J. (2023). Example citation text with hanging indent.");

        // Reset the indent for any subsequent paragraphs (optional).
        builder.ParagraphFormat.FirstLineIndent = 0;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "HangingIndent.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
