using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create a folder for all output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // ---------- Create the original document ----------
        var docOriginal = new Document();
        var builder = new DocumentBuilder(docOriginal);
        builder.Writeln("Hello world! This is the original document.");
        builder.Writeln("This paragraph will stay the same.");
        builder.Writeln("Original third paragraph.");

        // ---------- Create the edited document with intentional differences ----------
        var docEdited = new Document();
        builder = new DocumentBuilder(docEdited);
        builder.Writeln("Hello world! This is the edited document."); // changed text
        builder.Writeln("This paragraph will stay the same.");       // unchanged
        builder.Writeln("Edited third paragraph with extra text."); // changed text

        // Ensure both documents have no revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Perform the comparison. The author name and timestamp are required.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the document that now contains revisions.
        string comparedPath = Path.Combine(outputDir, "Compared.docx");
        docOriginal.Save(comparedPath);

        // ---------- Custom logger: write revision details to a CSV file ----------
        string logPath = Path.Combine(outputDir, "RevisionLog.csv");
        using var writer = new StreamWriter(logPath);
        writer.WriteLine("RevisionType,Author,DateTime"); // CSV header

        foreach (Revision rev in docOriginal.Revisions)
        {
            // ISO 8601 format for the timestamp.
            writer.WriteLine($"{rev.RevisionType},{rev.Author},{rev.DateTime:O}");
        }

        // The example finishes automatically; no user interaction is required.
    }
}
