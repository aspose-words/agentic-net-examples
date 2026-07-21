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

        // Write a sample citation paragraph.
        builder.Writeln("Doe, J. (2023). *Example Citation Title*. Publisher.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "HangingIndent.docx");
        doc.Save(outputPath);
    }
}
