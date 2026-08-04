using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a hanging indent of 0.25 inches (18 points). Negative value creates a hanging indent.
        builder.ParagraphFormat.FirstLineIndent = -18; // 0.25 inch = 0.25 * 72 points

        // Add a sample citation paragraph.
        builder.Writeln("Doe, J. (2020). Example citation text with a hanging indent for formatting purposes.");

        // Determine the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HangingIndent.docx");

        // Ensure the directory exists (in case the current directory is a root path).
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document.
        doc.Save(outputPath);
    }
}
