using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original legal contract document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("THIS AGREEMENT is made on the 1st day of January, 2023.");
        builder.Font.Bold = true; // Formatting that will be ignored.
        builder.Writeln("The Parties agree to the following terms and conditions.");
        builder.Font.Bold = false;
        builder.Writeln("Clause 1: Confidentiality.");
        builder.Writeln("Clause 2: Termination.");

        // Create the revised legal contract document with formatting changes and a textual change.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("THIS AGREEMENT is made on the 1st day of January, 2023."); // Same text, different formatting.
        builder2.Font.Bold = false; // Different formatting (not bold).
        builder2.Writeln("The Parties agree to the following terms and conditions."); // Same text.
        builder2.Font.Bold = false;
        builder2.Writeln("Clause 1: Confidentiality."); // Same text.
        builder2.Writeln("Clause 2: Termination and Severance."); // Textual change.

        // Configure comparison options to ignore formatting changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        original.Compare(revised, "LegalTeam", DateTime.Now, options);

        // Inspect and report revisions.
        int revisionCount = original.Revisions.Count;
        Console.WriteLine($"Total revisions after comparison (ignoring formatting): {revisionCount}");

        foreach (Revision rev in original.Revisions)
        {
            // Get the text of the revision from its parent node, if available.
            string text = rev.ParentNode?.GetText() ?? string.Empty;
            Console.WriteLine($"{rev.RevisionType}: \"{text.Trim()}\"");
        }

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LegalContractComparison.docx");
        original.Save(outputPath);
        Console.WriteLine($"Comparison document saved to: {outputPath}");
    }
}
