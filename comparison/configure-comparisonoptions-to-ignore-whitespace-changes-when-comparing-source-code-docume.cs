using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original documentation file.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("public class Sample");
        builderOriginal.Writeln("{");
        builderOriginal.Writeln("    // This method adds two numbers");
        builderOriginal.Writeln("    public int Add(int a, int b)");
        builderOriginal.Writeln("    {");
        builderOriginal.Writeln("        return a + b;");
        builderOriginal.Writeln("    }");
        builderOriginal.Writeln("}");

        // Create the revised documentation file with different whitespace.
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.Writeln("public class Sample");
        builderRevised.Writeln("{");
        builderRevised.Writeln("");
        builderRevised.Writeln("    // This method adds two numbers");
        builderRevised.Writeln("    public int Add( int a , int b )");
        builderRevised.Writeln("    {");
        builderRevised.Writeln("        return a + b ;");
        builderRevised.Writeln("    }");
        builderRevised.Writeln("}");

        // Configure comparison options to ignore formatting (including whitespace changes).
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            CompareMoves = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison.
        originalDoc.Compare(revisedDoc, "DocComparer", DateTime.Now, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        originalDoc.Save(resultPath);

        // Output the number of revisions detected (should be zero when whitespace is ignored).
        Console.WriteLine($"Revisions count after ignoring whitespace: {originalDoc.Revisions.Count}");
    }
}
