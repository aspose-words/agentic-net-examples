using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package reference
using Newtonsoft.Json;        // Required package reference

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // Create a sample document containing Unicode em dashes.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example—text with an em dash.");
        builder.Writeln("Another line—another dash.");
        doc.Save(inputPath);

        // -------------------------------------------------
        // Load the document we just created.
        // -------------------------------------------------
        Document loaded = new Document(inputPath);

        // -------------------------------------------------
        // Define a regular expression that matches the Unicode em dash (U+2014).
        // -------------------------------------------------
        Regex emDashRegex = new Regex("\u2014");

        // -------------------------------------------------
        // Replace each em dash with a standard hyphen.
        // -------------------------------------------------
        int replacedCount = loaded.Range.Replace(emDashRegex, "-", new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one em dash replacement.");

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        loaded.Save(outputPath);
    }
}
