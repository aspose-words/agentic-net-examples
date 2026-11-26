using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly
using Newtonsoft.Json; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple consecutive spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is   a    test     document.");
        builder.Writeln("Another    line   with  spaces.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches two or more spaces.
        Regex spaceRegex = new Regex(@" {2,}");

        // Replace each occurrence with a single space.
        int replacedCount = loaded.Range.Replace(spaceRegex, " ", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
