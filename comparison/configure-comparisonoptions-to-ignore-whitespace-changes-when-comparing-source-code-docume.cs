using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonExample");
        Directory.CreateDirectory(workDir);

        // Create the original documentation file with extra whitespace.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("public class Sample");
        builderOriginal.Writeln("{");
        builderOriginal.Writeln("    // This method does something");
        builderOriginal.Writeln("    public void DoWork( )   ");
        builderOriginal.Writeln("    {");
        builderOriginal.Writeln("        // TODO: implement");
        builderOriginal.Writeln("    }");
        builderOriginal.Writeln("}");
        string originalPath = Path.Combine(workDir, "Original.docx");
        originalDoc.Save(originalPath);

        // Create the revised documentation file with trimmed whitespace.
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.Writeln("public class Sample");
        builderRevised.Writeln("{");
        builderRevised.Writeln("// This method does something");
        builderRevised.Writeln("public void DoWork()");
        builderRevised.Writeln("{");
        builderRevised.Writeln("// TODO: implement");
        builderRevised.Writeln("}");
        builderRevised.Writeln("}");
        string revisedPath = Path.Combine(workDir, "Revised.docx");
        revisedDoc.Save(revisedPath);

        // Load the documents back (simulating real file usage).
        Document doc1 = new Document(originalPath);
        Document doc2 = new Document(revisedPath);

        // Configure compare options to ignore whitespace/formatting changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        doc1.Compare(doc2, "Comparer", DateTime.Now, options);

        // Verify that whitespace changes were ignored (no revisions expected).
        int revisionCount = doc1.Revisions.Count;
        Console.WriteLine($"Revisions count after comparison (ignoring whitespace): {revisionCount}");

        // Save the comparison result.
        string resultPath = Path.Combine(workDir, "ComparisonResult.docx");
        doc1.Save(resultPath);
    }
}
