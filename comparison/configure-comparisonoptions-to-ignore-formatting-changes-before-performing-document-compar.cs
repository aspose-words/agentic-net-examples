using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the first version of the legal contract.
        Document contractV1 = new Document();
        DocumentBuilder builderV1 = new DocumentBuilder(contractV1);
        builderV1.Font.Size = 12;
        builderV1.Font.Bold = false;
        builderV1.Writeln("CONFIDENTIAL LEGAL AGREEMENT");
        builderV1.Writeln("This agreement is made between Party A and Party B.");
        builderV1.Writeln("All terms and conditions are binding.");

        // Create the second version with different formatting only.
        Document contractV2 = new Document();
        DocumentBuilder builderV2 = new DocumentBuilder(contractV2);
        builderV2.Font.Size = 14;               // Different font size
        builderV2.Font.Bold = true;             // Different bold setting
        builderV2.Writeln("CONFIDENTIAL LEGAL AGREEMENT");
        builderV2.Writeln("This agreement is made between Party A and Party B.");
        builderV2.Writeln("All terms and conditions are binding.");

        // Configure comparison options to ignore formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        contractV1.Compare(contractV2, "LegalTeam", DateTime.Now, compareOptions);

        // Verify that no revisions were generated because formatting differences are ignored.
        if (contractV1.Revisions.Count != 0)
        {
            throw new InvalidOperationException(
                $"Expected zero revisions, but found {contractV1.Revisions.Count}.");
        }

        // Save the comparison result (should be identical to the original content).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        contractV1.Save(outputPath);
    }
}
