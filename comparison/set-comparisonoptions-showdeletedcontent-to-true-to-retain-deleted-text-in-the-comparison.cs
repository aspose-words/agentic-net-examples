using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some text, including a line that will be deleted.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is a sample paragraph.");
        builderOriginal.Writeln("This line will be deleted.");
        builderOriginal.Writeln("Another line that stays.");

        // Create the revised document where the line to be deleted is omitted.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is a sample paragraph.");
        builderRevised.Writeln("Another line that stays.");

        // Configure compare options. The ShowDeletedContent property is not required
        // because deleted text is retained by default in the comparison output.
        CompareOptions options = new CompareOptions();

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
    }
}
