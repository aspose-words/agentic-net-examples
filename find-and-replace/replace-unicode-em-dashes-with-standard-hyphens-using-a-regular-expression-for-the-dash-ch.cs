using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package as per task specification
using Newtonsoft.Json;        // Required package as per task specification

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document containing Unicode em dashes (U+2014).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Sample text with em dashes.
        builder.Writeln("This—sentence—contains—em—dashes.");
        builder.Writeln("Another line—here.");
        // Save the source document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and perform the replacement using a regex.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regex that matches the Unicode em dash character.
        Regex emDashRegex = new Regex("\u2014"); // U+2014

        // Replace each em dash with a standard hyphen.
        int replacedCount = loaded.Range.Replace(emDashRegex, "-", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one em dash replacement, but none were found.");

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // Optional: Output a simple confirmation (no interactive prompts).
        Console.WriteLine($"Replaced {replacedCount} em dash(es). Output saved to: {outputPath}");
    }
}
