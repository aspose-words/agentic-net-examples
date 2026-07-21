using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the original legal contract document.
        Document originalContract = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalContract);
        builderOriginal.Writeln("CONFIDENTIAL AGREEMENT");
        builderOriginal.Font.Size = 12;
        builderOriginal.Font.Bold = true; // Formatting that will be ignored.
        builderOriginal.Writeln("This agreement is made between Party A and Party B.");
        builderOriginal.Font.Bold = false;
        builderOriginal.Writeln("Terms and conditions apply.");

        // Create the revised legal contract document with formatting changes.
        Document revisedContract = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedContract);
        builderRevised.Writeln("CONFIDENTIAL AGREEMENT");
        builderRevised.Font.Size = 12;
        builderRevised.Font.Italic = true; // Different formatting.
        builderRevised.Writeln("This agreement is made between Party A and Party B.");
        builderRevised.Font.Italic = false;
        builderRevised.Writeln("Terms and conditions apply.");

        // Configure comparison options to ignore formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,
            // Other flags remain at their defaults (false).
        };

        // Perform the comparison. The original document will receive revisions.
        originalContract.Compare(revisedContract, "LegalTeam", DateTime.Now, compareOptions);

        // Verify that revisions exist (text differences are none, only formatting ignored).
        int revisionCount = originalContract.Revisions.Count;
        Console.WriteLine($"Number of revisions after comparison (should be 0 if only formatting differed): {revisionCount}");

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        originalContract.Save(outputPath);
        Console.WriteLine($"Comparison result saved to: {outputPath}");
    }
}
