using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing Unicode em dashes.
        var inputPath = "input.docx";
        var outputPath = "output.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example sentence—containing an em dash.");
        builder.Writeln("Another line—here as well.");
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Define a regex that matches the Unicode em dash (U+2014).
        var emDashRegex = new Regex("\u2014");

        // Perform the replacement with a standard hyphen.
        var options = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace(emDashRegex, "-", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one em dash replacement.");

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }
}
