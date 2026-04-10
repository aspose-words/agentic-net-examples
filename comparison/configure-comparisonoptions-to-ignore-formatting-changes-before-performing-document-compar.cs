using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class CompareIgnoreFormattingExample
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create the original legal contract document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("CONFIDENTIAL AGREEMENT");
        builder.Font.Size = 12;
        builder.Font.Bold = true;
        builder.Writeln("This agreement is made between Party A and Party B.");
        builder.Font.Bold = false;
        builder.Writeln("The term of this agreement shall be five (5) years.");

        // Clone the original and introduce some changes, including formatting changes.
        Document revised = (Document)original.Clone(true);
        DocumentBuilder revBuilder = new DocumentBuilder(revised);
        // Change some text.
        Paragraph firstParagraph = revised.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "CONFIDENTIAL CONTRACT";
        // Change formatting (which we will ignore).
        revBuilder.MoveTo(firstParagraph);
        revBuilder.Font.Italic = true;
        revBuilder.Writeln("Additional clause added for clarification.");

        // Configure comparison options to ignore formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,
            // Keep other flags at default (false) to capture text changes.
            CompareMoves = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison.
        original.Compare(revised, "LegalTeam", DateTime.Now, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(artifactsDir, "ContractComparison.docx");
        original.Save(resultPath);

        // Output revision count to verify that formatting changes were ignored.
        Console.WriteLine($"Total revisions after comparison (formatting ignored): {original.Revisions.Count}");
    }
}
