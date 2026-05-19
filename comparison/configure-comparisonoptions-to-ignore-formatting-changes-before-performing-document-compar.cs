using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create the original contract document.
        // -------------------------------------------------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("CONFIDENTIAL AGREEMENT");
        builder.Font.Size = 14;
        builder.Font.Bold = true;
        builder.Writeln("This contract is between Party A and Party B.");
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("All terms are subject to change.");

        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "OriginalContract.docx");
        original.Save(originalPath);

        // -------------------------------------------------
        // Create the revised contract with formatting changes.
        // -------------------------------------------------
        Document revised = new Document();
        DocumentBuilder revBuilder = new DocumentBuilder(revised);
        revBuilder.Writeln("CONFIDENTIAL AGREEMENT");
        revBuilder.Font.Size = 12;               // Different font size.
        revBuilder.Font.Bold = false;            // No bold.
        revBuilder.Writeln("This contract is between Party A and Party B.");
        revBuilder.Font.Bold = true;             // Now bold.
        revBuilder.Font.Italic = false;
        revBuilder.Writeln("All terms are subject to change.");
        // Additional content (no underline needed for this example).
        revBuilder.Writeln("Additional clause added.");

        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisedContract.docx");
        revised.Save(revisedPath);

        // -------------------------------------------------
        // Configure comparison to ignore formatting changes.
        // -------------------------------------------------
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison. Only content changes will generate revisions.
        original.Compare(revised, "LegalTeam", DateTime.Now, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ContractComparison.docx");
        original.Save(resultPath);

        // Output the number of revisions detected (formatting-only changes are ignored).
        Console.WriteLine($"Revisions count (ignoring formatting): {original.Revisions.Count}");
    }
}
