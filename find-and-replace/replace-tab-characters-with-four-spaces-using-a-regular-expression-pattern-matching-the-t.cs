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
        builder.Writeln("Item1\tItem2\tItem3");
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Define a regex that matches a tab character.
        Regex tabRegex = new Regex(@"\t");

        // Replace each tab with four spaces.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacementCount = loaded.Range.Replace(tabRegex, "    ", options);

        // Ensure that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one tab replacement.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loaded.Save(outputPath);
    }
}
