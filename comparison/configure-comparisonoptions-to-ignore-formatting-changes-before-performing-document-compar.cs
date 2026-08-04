using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original contract document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("CONFIDENTIAL AGREEMENT");
        builderOriginal.Font.Bold = true;
        builderOriginal.Writeln("This agreement is made between Party A and Party B.");
        builderOriginal.Font.Bold = false;
        builderOriginal.Font.Italic = true;
        builderOriginal.Writeln("The term of this agreement shall be five (5) years.");
        builderOriginal.Font.Italic = false;

        // Create the revised contract document with the same text but different formatting.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("CONFIDENTIAL AGREEMENT");
        builderRevised.Font.Bold = false; // different formatting
        builderRevised.Writeln("This agreement is made between Party A and Party B.");
        builderRevised.Font.Bold = true;
        builderRevised.Font.Italic = false; // different formatting
        builderRevised.Writeln("The term of this agreement shall be five (5) years.");
        builderRevised.Font.Italic = true;

        // Configure comparison options to ignore formatting changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        original.Compare(revised, "LegalTeam", DateTime.Now, options);

        // Verify that no revisions were generated because only formatting differences exist.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException($"Expected zero revisions, but found {original.Revisions.Count}.");
        }

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ContractComparisonResult.docx");
        original.Save(outputPath);
    }
}
