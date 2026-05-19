using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Create first documentation file with specific whitespace.
        Document docOriginal = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(docOriginal);
        builderOriginal.Writeln("public class Example");
        builderOriginal.Writeln("{");
        builderOriginal.Writeln("    // Method");
        builderOriginal.Writeln("    public void DoWork()");
        builderOriginal.Writeln("    {");
        builderOriginal.Writeln("        Console.WriteLine(\"Hello\");");
        builderOriginal.Writeln("    }");
        builderOriginal.Writeln("}");

        // Create second documentation file with different whitespace (extra spaces).
        Document docRevised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(docRevised);
        builderRevised.Writeln("public class Example");
        builderRevised.Writeln("{");
        builderRevised.Writeln("\t// Method"); // Tab instead of spaces
        builderRevised.Writeln("\tpublic void DoWork()");
        builderRevised.Writeln("\t{");
        builderRevised.Writeln("\t    Console.WriteLine(\"Hello\");"); // Different indentation
        builderRevised.Writeln("\t}");
        builderRevised.Writeln("}");

        // Configure compare options to ignore formatting (including whitespace) changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform comparison.
        docOriginal.Compare(docRevised, "SourceCodeComparer", DateTime.Now, compareOptions);

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        docOriginal.Save(outputPath);

        // Output revision count to console.
        int revisionCount = docOriginal.Revisions.Count;
        Console.WriteLine($"Comparison completed. Revisions detected (excluding whitespace): {revisionCount}");
    }
}
