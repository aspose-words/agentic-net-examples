using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly but ensures reference

public class Program
{
    public static void Main()
    {
        // Create a sample document with bullet characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("• Item 1");
        builder.Writeln("• Item 2");
        builder.Writeln("• Item 3");
        builder.Writeln("Regular paragraph without bullet.");

        // Save the sample input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the bullet character (U+2022).
        Regex bulletRegex = new Regex("\u2022");

        // Replace the bullet character with a different bullet style (U+25E6).
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(bulletRegex, "\u25E6", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one bullet replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacedCount} bullet character(s). Output saved to '{outputPath}'.");
    }
}
