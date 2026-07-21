using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with bullet characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Shopping List:");
        builder.Writeln("\u2022 Apples");
        builder.Writeln("\u2022 Bananas");
        builder.Writeln("\u2022 Oranges");
        string inputPath = Path.Combine(outputDir, "input.docx");
        doc.Save(inputPath);

        // Load the document (could also continue using the same instance).
        Document loaded = new Document(inputPath);

        // Define a regex that matches the bullet character (U+2022).
        Regex bulletRegex = new Regex("\u2022");

        // Replace the bullet with a different style (U+25E6 – white bullet).
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(bulletRegex, "\u25E6", options);

        // Ensure at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No bullet characters were replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "output.docx");
        loaded.Save(outputPath);
    }
}
