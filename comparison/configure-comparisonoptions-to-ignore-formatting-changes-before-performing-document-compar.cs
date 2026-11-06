using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the first version of a legal contract.
        Document contractV1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(contractV1);
        builder1.Font.Size = 12;
        builder1.Font.Name = "Times New Roman";
        builder1.Writeln("CONFIDENTIAL LEGAL CONTRACT");
        builder1.Font.Bold = true;
        builder1.Writeln("This agreement is made on the 1st of January 2023.");
        builder1.Font.Bold = false;
        builder1.Writeln("The parties agree to the following terms and conditions...");
        builder1.Font.Italic = true;
        builder1.Writeln("Section 1: Definitions");
        builder1.Font.Italic = false;

        // Create the second version with text changes and formatting changes.
        Document contractV2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(contractV2);
        builder2.Font.Size = 14; // Different size (formatting)
        builder2.Font.Name = "Arial"; // Different font (formatting)
        builder2.Writeln("CONFIDENTIAL LEGAL CONTRACT");
        builder2.Font.Bold = true;
        builder2.Writeln("This agreement is made on the 2nd of February 2023."); // Text change
        builder2.Font.Bold = false;
        builder2.Writeln("The parties agree to the following terms and conditions...");
        builder2.Font.Italic = true;
        builder2.Writeln("Section 1: Definitions and Interpretations"); // Text change
        builder2.Font.Italic = false;

        // Configure comparison to ignore formatting changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        contractV1.Compare(contractV2, "LegalTeam", DateTime.Now, options);

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "contract-comparison.docx");
        contractV1.Save(outputPath);

        // Report the number of revisions detected (formatting changes are ignored).
        Console.WriteLine($"Total revisions detected (ignoring formatting): {contractV1.Revisions.Count}");
    }
}
