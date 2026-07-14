using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create the original document with some code and whitespace.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("public class Sample");
        builderOriginal.Writeln("{");
        builderOriginal.Writeln("    public void Method()");
        builderOriginal.Writeln("    {");
        builderOriginal.Writeln("        int x = 1;");
        builderOriginal.Writeln("    }");
        builderOriginal.Writeln("}");

        // Create the revised document that differs only by whitespace (extra spaces and blank lines).
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("public class Sample");
        builderRevised.Writeln("{");
        builderRevised.Writeln(""); // blank line
        builderRevised.Writeln("    public void Method()");
        builderRevised.Writeln("    {");
        builderRevised.Writeln("        int  x = 1;   "); // extra spaces
        builderRevised.Writeln("    }");
        builderRevised.Writeln("}");

        // Normalize whitespace in both documents so that only meaningful text differences remain.
        NormalizeDocumentWhitespace(original);
        NormalizeDocumentWhitespace(revised);

        // Configure compare options (ignore formatting changes that are not relevant to source code).
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Verify that no revisions were generated because only whitespace changes existed.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException(
                $"Expected zero revisions, but found {original.Revisions.Count}.");
        }

        // Save the result document (it will be identical to the original after normalization).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath, SaveFormat.Docx);

        Console.WriteLine($"Comparison completed. No revisions detected. Result saved to: {outputPath}");
    }

    /// <summary>
    /// Removes insignificant whitespace from all paragraphs in the document:
    /// - Trims leading/trailing spaces.
    /// - Replaces multiple consecutive spaces with a single space.
    /// - Removes empty paragraphs.
    /// </summary>
    private static void NormalizeDocumentWhitespace(Document doc)
    {
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>()
                           .ToList();

        foreach (var paragraph in paragraphs)
        {
            string text = paragraph.GetText(); // Includes paragraph break at the end.
            // Remove the paragraph break for processing.
            if (text.EndsWith("\r") || text.EndsWith("\n"))
                text = text.TrimEnd('\r', '\n');

            // Normalize spaces.
            string normalized = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();

            if (string.IsNullOrEmpty(normalized))
            {
                // Remove empty paragraphs.
                paragraph.Remove();
                continue;
            }

            // Clear existing runs and add a single run with normalized text.
            paragraph.Runs.Clear();
            paragraph.AppendChild(new Run(doc, normalized));
        }
    }
}
