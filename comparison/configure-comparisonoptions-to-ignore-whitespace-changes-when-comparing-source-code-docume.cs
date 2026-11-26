using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original documentation file.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("/// <summary>");
        builder1.Writeln("/// This method does something.");
        builder1.Writeln("/// </summary>");
        builder1.Writeln("public void DoWork(){ }");

        // Create the revised documentation file with whitespace changes only.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("/// <summary>   "); // trailing spaces
        builder2.Writeln("///   This method does something."); // leading spaces
        builder2.Writeln("/// </summary>");
        builder2.Writeln(""); // extra blank line
        builder2.Writeln("public void DoWork(){ }");

        // Configure compare options to ignore formatting (including whitespace).
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Verify that whitespace changes were ignored (no revisions expected).
        int revisionCount = original.Revisions.Count;

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);

        // Output the result information.
        Console.WriteLine($"Revisions count: {revisionCount}");
        Console.WriteLine($"Comparison result saved to: {outputPath}");
    }
}
