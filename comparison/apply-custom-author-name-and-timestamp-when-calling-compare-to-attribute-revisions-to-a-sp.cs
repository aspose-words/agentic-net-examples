using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Define a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original paragraph.");
        builder.Writeln("It contains some sample text.");

        // Create the edited document with intentional differences.
        Document edited = new Document();
        DocumentBuilder editedBuilder = new DocumentBuilder(edited);
        editedBuilder.Writeln("This is the edited paragraph."); // Changed text.
        editedBuilder.Writeln("It contains some sample text."); // Same as original.
        editedBuilder.Writeln("An additional line was added.");   // New line.

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count == 0 && edited.Revisions.Count == 0)
        {
            // Apply custom author name and timestamp for the revisions.
            string customAuthor = "JaneDoe";
            DateTime customDate = new DateTime(2023, 12, 31, 23, 59, 59);
            original.Compare(edited, customAuthor, customDate);
        }

        // Output revision details to the console.
        Console.WriteLine($"Total revisions created: {original.Revisions.Count}");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"Revision type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
            Console.WriteLine($"Changed text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document that now contains the revisions.
        string outputPath = Path.Combine(outputDir, "ComparisonResult.docx");
        original.Save(outputPath);
        Console.WriteLine($"Comparison document saved to: {outputPath}");
    }
}
