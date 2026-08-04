using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Create a sample document containing tab characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Item\tQuantity\tPrice");
        builder.Writeln("Apple\t10\t$1.00");
        builder.Writeln("Banana\t5\t$0.50");
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Define a regex that matches a tab character.
        Regex tabRegex = new Regex("\t");

        // Replace each tab with four spaces.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(tabRegex, "    ", options);

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one tab replacement.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
