using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json; // Included as per required packages

public class Program
{
    public static void Main()
    {
        // Create a sample document containing tab characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Item1\tItem2\tItem3");
        builder.Writeln("ValueA\tValueB\tValueC");
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that matches a tab character.
        Regex tabRegex = new Regex(@"\t");

        // Replace each tab with four spaces.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(tabRegex, "    ", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one tab character to be replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loaded.Save(outputPath);
    }
}
