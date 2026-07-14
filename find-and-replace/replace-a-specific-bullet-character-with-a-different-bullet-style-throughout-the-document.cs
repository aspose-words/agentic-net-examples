using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document containing bullet characters.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several lines that start with the bullet character "•".
        builder.Writeln("• First item");
        builder.Writeln("• Second item");
        builder.Writeln("• Third item");
        builder.Writeln("Regular paragraph without bullet.");

        // Save the document so it can be loaded later.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and replace the bullet character with a new one.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regular expression that matches the bullet character "•".
        Regex bulletRegex = new Regex("•");

        // Replacement bullet character, e.g., "◦".
        const string newBullet = "◦";

        // Perform the replacement across the whole document.
        int replacedCount = loaded.Range.Replace(bulletRegex, newBullet, new FindReplaceOptions());

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one bullet character to be replaced.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
