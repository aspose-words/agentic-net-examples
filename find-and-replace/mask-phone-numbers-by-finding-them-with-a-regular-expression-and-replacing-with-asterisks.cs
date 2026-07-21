using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // Create a sample document containing various phone number formats.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Customer contacts:");
        builder.Writeln("John Doe: 123-456-7890");
        builder.Writeln("Jane Smith: (123) 456 7890");
        builder.Writeln("Bob Johnson: 1234567890");
        builder.Writeln("End of list.");
        doc.Save(inputPath);

        // ---------------------------------------------------------------
        // Load the document, mask phone numbers using a regular expression.
        // ---------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regex matches common US phone number patterns.
        Regex phoneRegex = new Regex(@"\b(?:\d{3}[-.\s]?\d{3}[-.\s]?\d{4}|\(\d{3}\)\s?\d{3}[-.\s]?\d{4})\b");

        // Replace each match with ten asterisks (preserving length for privacy).
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(phoneRegex, "**********", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No phone numbers were found to mask.");

        // Save the masked document.
        loaded.Save(outputPath);

        // Optional: write a brief confirmation to the console.
        Console.WriteLine($"Phone numbers masked: {replacedCount}");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}
