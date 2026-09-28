using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with dates in MM-DD-YYYY format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The first date is 12-31-2023.");
        builder.Writeln("Another date: 01-15-2024.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Regular expression to match dates in MM-DD-YYYY format.
        Regex dateRegex = new Regex(@"\b(\d{2})-(\d{2})-(\d{4})\b");

        // Replacement pattern to convert to YYYY-MM-DD format.
        const string replacementPattern = "$3-$1-$2";

        // Perform the regex replace across the document.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace(dateRegex, replacementPattern, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
