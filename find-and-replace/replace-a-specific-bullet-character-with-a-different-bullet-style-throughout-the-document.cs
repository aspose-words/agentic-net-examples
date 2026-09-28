using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package reference
using Newtonsoft.Json; // Required package reference

public class Program
{
    public static void Main()
    {
        // Create a sample document with bullet characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("• First item");
        builder.Writeln("• Second item");
        builder.Writeln("– Not a bullet to replace");
        builder.Writeln("• Third item");
        // Save the original document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the specific bullet character (U+2022).
        Regex bulletRegex = new Regex(@"\u2022");

        // Replace the bullet character with a different bullet style (U+25E6).
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(bulletRegex, "◦", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one bullet replacement, but none were made.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: Write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacedCount} bullet character(s). Output saved to '{outputPath}'.");
    }
}
