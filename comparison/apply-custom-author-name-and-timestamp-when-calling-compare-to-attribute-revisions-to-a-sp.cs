using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Define custom author name and timestamp for the comparison.
        const string customAuthor = "JaneDoe";
        DateTime customDate = new DateTime(2023, 5, 1, 10, 30, 0, DateTimeKind.Utc);

        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");
        builderOriginal.Writeln("It contains a line that will be changed.");

        // Create the revised document with intentional differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the original paragraph."); // unchanged line.
        builderRevised.Writeln("It contains a line that has been edited."); // edited line.

        // Perform the comparison, attributing all revisions to the custom author and timestamp.
        original.Compare(revised, customAuthor, customDate);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");

        // Ensure every revision has the expected author and timestamp.
        foreach (Revision rev in original.Revisions)
        {
            if (!string.Equals(rev.Author, customAuthor, StringComparison.Ordinal))
                throw new InvalidOperationException($"Revision author mismatch. Expected '{customAuthor}', got '{rev.Author}'.");

            // Compare only the date component to avoid minor differences in ticks.
            if (rev.DateTime.Date != customDate.Date)
                throw new InvalidOperationException($"Revision date mismatch. Expected '{customDate.Date:d}', got '{rev.DateTime.Date:d}'.");
        }

        // Save the compared document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);

        // Optional console output to indicate success.
        Console.WriteLine($"Comparison completed. Revisions attributed to '{customAuthor}' on {customDate:d}.");
        Console.WriteLine($"Result saved to: {outputPath}");
    }
}
