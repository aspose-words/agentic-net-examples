using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the first documentation file with specific whitespace.
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("/// <summary>");
        builderOriginal.Writeln("/// This method does something.");
        builderOriginal.Writeln("/// </summary>");
        builderOriginal.Writeln("public void DoWork()");
        builderOriginal.Writeln("{");
        builderOriginal.Writeln("    // Implementation");
        builderOriginal.Writeln("}");

        // Create the second documentation file that differs only by whitespace.
        Document docRevised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(docRevised);
        // Add extra spaces and blank lines.
        builderRevised.Writeln("/// <summary>");
        builderRevised.Writeln("");
        builderRevised.Writeln("///   This method does something.   ");
        builderRevised.Writeln("/// </summary>");
        builderRevised.Writeln("public void DoWork()");
        builderRevised.Writeln("{");
        builderRevised.Writeln("        // Implementation");
        builderRevised.Writeln("}");

        // Configure compare options to ignore formatting (including whitespace) changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now, compareOptions);

        // Output the number of revisions detected.
        Console.WriteLine($"Revisions detected: {docOriginal.Revisions.Count}");

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        docOriginal.Save(outputPath);
        Console.WriteLine($"Comparison document saved to: {outputPath}");
    }
}
