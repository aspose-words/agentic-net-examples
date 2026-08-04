using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package
using Newtonsoft.Json;        // Required package

public class Program
{
    public static void Main()
    {
        // Create a sample document with bullet characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("• First item");
        builder.Writeln("• Second item");
        builder.Writeln("• Third item");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Define a regular expression that matches the bullet character (U+2022).
        Regex bulletRegex = new Regex("\u2022");

        // Replace the bullet with an alternative bullet style (U+25E6).
        int replacedCount = loaded.Range.Replace(bulletRegex, "\u25E6", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one bullet replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }
}
