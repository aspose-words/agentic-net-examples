using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original legal contract document.
        Document originalContract = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalContract);
        // Title with bold formatting.
        builderOriginal.Font.Size = 16;
        builderOriginal.Font.Bold = true;
        builderOriginal.Writeln("CONFIDENTIAL AGREEMENT");
        // Reset formatting for body text.
        builderOriginal.Font.Size = 12;
        builderOriginal.Font.Bold = false;
        builderOriginal.Writeln("This Agreement is made between Party A and Party B.");
        builderOriginal.Writeln("The term of this Agreement shall be five (5) years.");
        // Save the original for reference (optional).
        originalContract.Save(Path.Combine(outputDir, "OriginalContract.docx"));

        // Create the revised legal contract document with some formatting changes.
        Document revisedContract = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedContract);
        // Title without bold formatting (formatting change we want to ignore).
        builderRevised.Font.Size = 16;
        builderRevised.Font.Bold = false; // Different formatting.
        builderRevised.Writeln("CONFIDENTIAL AGREEMENT");
        // Body text with a minor content change.
        builderRevised.Font.Size = 12;
        builderRevised.Font.Bold = false;
        builderRevised.Writeln("This Agreement is made between Party A and Party B.");
        builderRevised.Writeln("The term of this Agreement shall be six (6) years."); // Content change.
        revisedContract.Save(Path.Combine(outputDir, "RevisedContract.docx"));

        // Configure comparison options to ignore formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true // Ignore all formatting differences.
        };

        // Perform the comparison. The original document will receive revisions.
        originalContract.Compare(revisedContract, "LegalTeam", DateTime.Now, compareOptions);

        // Verify that revisions were created (there should be at least one due to content change).
        int revisionCount = originalContract.Revisions.Count;
        Console.WriteLine($"Revisions detected: {revisionCount}");

        // Save the comparison result.
        string resultPath = Path.Combine(outputDir, "ComparisonResult.docx");
        originalContract.Save(resultPath);
        Console.WriteLine($"Comparison document saved to: {resultPath}");
    }
}
