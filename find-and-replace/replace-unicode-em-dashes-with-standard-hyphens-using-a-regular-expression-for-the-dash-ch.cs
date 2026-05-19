using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing; // Required for FindReplaceOptions
using Newtonsoft.Json; // Kept for compatibility as in original example

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -------------------------------------------------
        // 1. Create a sample document containing em dashes.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example—text with an em dash.");
        builder.Writeln("Another line—showing multiple—em dashes.");
        doc.Save(inputPath);

        // -------------------------------------------------
        // 2. Load the document and replace each em dash with a hyphen.
        // -------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regular expression that matches the Unicode em dash character (U+2014).
        Regex emDashRegex = new Regex("\u2014");

        // Perform the replacement using FindReplaceOptions.
        int replacedCount = loaded.Range.Replace(emDashRegex, "-", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No em dash characters were found for replacement.");

        // -------------------------------------------------
        // 3. Save the modified document.
        // -------------------------------------------------
        loaded.Save(outputPath);
    }
}
