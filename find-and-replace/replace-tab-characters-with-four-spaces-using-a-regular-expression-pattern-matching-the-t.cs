using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing tab characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line with\ttab.");
        builder.Writeln("Second line with\ttwo tabs:\t\tend.");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches a tab character.
        Regex tabRegex = new Regex(@"\t");

        // Replace each tab with four spaces.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(tabRegex, "    ", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No tab characters were replaced.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replaced {replacedCount} tab character(s). Output saved to '{outputPath}'.");
    }
}
