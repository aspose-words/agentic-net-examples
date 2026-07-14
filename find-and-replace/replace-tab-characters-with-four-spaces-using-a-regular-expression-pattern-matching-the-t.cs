using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package, not used directly in this example
using Newtonsoft.Json;        // Required package, not used directly in this example

public class Program
{
    public static void Main()
    {
        // Create a sample document containing tab characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Item1\tItem2\tItem3");
        builder.Writeln("ValueA\tValueB\tValueC");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches a tab character.
        Regex tabRegex = new Regex(@"\t");

        // Replace each tab with four spaces.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacementCount = loaded.Range.Replace(tabRegex, "    ", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one tab replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: indicate success (no interactive input required).
        Console.WriteLine($"Replaced {replacementCount} tab(s). Output saved to '{outputPath}'.");
    }
}
