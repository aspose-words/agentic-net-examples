using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing the default bullet character (U+2022).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("\u2022 First item");
        builder.Writeln("\u2022 Second item");
        builder.Writeln("\u2022 Third item");

        // Save the source document (optional, demonstrates file I/O).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the saved file.
        Document loaded = new Document(inputPath);

        // Regular expression that matches the bullet character.
        Regex bulletRegex = new Regex("\u2022");

        // New bullet character to replace with (e.g., white bullet U+25E6).
        const string newBullet = "\u25E6";

        // Perform the replacement throughout the document.
        int replacedCount = loaded.Range.Replace(bulletRegex, newBullet, new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one bullet character to be replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
