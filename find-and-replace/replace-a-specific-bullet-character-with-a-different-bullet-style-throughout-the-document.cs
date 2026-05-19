using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with bullet characters.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("• First item");
        builder.Writeln("• Second item");
        builder.Writeln("Regular paragraph without bullet.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        var loadedDoc = new Document(inputPath);

        // Define a regular expression that matches the bullet character (U+2022).
        var bulletRegex = new Regex("\u2022");

        // Replace the bullet with an alternative bullet style (U+25E6).
        int replacedCount = loadedDoc.Range.Replace(bulletRegex, "\u25E6", new FindReplaceOptions());

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one bullet replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
