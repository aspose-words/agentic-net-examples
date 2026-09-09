using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not directly used in this example
using Newtonsoft.Json; // Required package for JSON report

public class BulletReplaceExample
{
    public static void Main()
    {
        // Define file names in the current directory
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "report.json");

        // -----------------------------------------------------------------
        // 1. Create a sample document containing the original bullet character (U+2022)
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a simple bullet list using the standard bullet character "•"
        builder.Writeln("• Item 1");
        builder.Writeln("• Item 2");
        builder.Writeln("Regular paragraph without bullet.");
        builder.Writeln("• Item 3");

        // Save the source document
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and replace the bullet character with a new one (U+25E6)
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regular expression that matches the bullet character "•"
        Regex bulletRegex = new Regex("\u2022");

        // Replacement bullet character "◦"
        string newBullet = "\u25E6";

        // Perform the replacement across the whole document
        int replacedCount = loaded.Range.Replace(bulletRegex, newBullet, new FindReplaceOptions());

        // Validate that at least one replacement occurred
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one bullet character to be replaced.");

        // Save the modified document
        loaded.Save(outputPath);

        // -----------------------------------------------------------------
        // 3. Write a simple JSON report containing the number of replacements
        // -----------------------------------------------------------------
        var report = new { Replacements = replacedCount };
        File.WriteAllText(reportPath, JsonConvert.SerializeObject(report, Formatting.Indented));

        // Optional console output (does not require user interaction)
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
