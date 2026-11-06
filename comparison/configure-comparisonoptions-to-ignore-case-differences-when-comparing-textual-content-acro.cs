using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with mixed‑case text.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello World!");

        // Create the revised document with the same text but different case.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("hello world!");

        // Configure comparison to ignore case differences.
        CompareOptions options = new CompareOptions();
        options.IgnoreCaseChanges = true;

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Verify that no revisions were generated because case differences are ignored.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Revisions were created despite ignoring case differences.");
        }

        // Save the resulting document.
        original.Save("case-ignored-comparison.docx");
    }
}
