using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some source code.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("public class Sample");
        builderOriginal.Writeln("{");
        builderOriginal.Writeln("    public void Method()");
        builderOriginal.Writeln("    {");
        builderOriginal.Writeln("        int x = 1;");
        builderOriginal.Writeln("    }");
        builderOriginal.Writeln("}");

        // Create the revised document that differs only by whitespace.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("public class Sample");
        builderRevised.Writeln("{");
        builderRevised.Writeln("    public void Method( )"); // added space before parentheses
        builderRevised.Writeln("    {");
        builderRevised.Writeln("        int   x = 1 ;"); // extra spaces
        builderRevised.Writeln("    }");
        builderRevised.Writeln("}");

        // Save the sample documents (optional, for inspection).
        string basePath = Directory.GetCurrentDirectory();
        string originalPath = Path.Combine(basePath, "original.docx");
        string revisedPath = Path.Combine(basePath, "revised.docx");
        original.Save(originalPath);
        revised.Save(revisedPath);

        // Configure comparison options to ignore whitespace (treated as formatting) changes.
        CompareOptions options = new CompareOptions();
        options.IgnoreFormatting = true; // Whitespace changes are considered formatting.

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Verify that no revisions were generated because whitespace changes are ignored.
        int revisionCount = original.Revisions.Count;
        Console.WriteLine($"Revisions count after ignoring whitespace: {revisionCount}");

        // Save the comparison result.
        string resultPath = Path.Combine(basePath, "compare-whitespace.docx");
        original.Save(resultPath);
    }
}
