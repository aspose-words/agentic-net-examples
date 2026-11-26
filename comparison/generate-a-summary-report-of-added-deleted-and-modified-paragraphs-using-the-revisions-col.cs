using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1: unchanged.");
        builder.Writeln("Paragraph 2: to be deleted.");
        builder.Writeln("Paragraph 3: to be modified.");

        // Create the revised document with changes:
        // - Paragraph 2 is omitted (deleted).
        // - Paragraph 3 text is changed (modified).
        // - A new paragraph is added.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Paragraph 1: unchanged.");
        builder2.Writeln("Paragraph 3: modified content.");
        builder2.Writeln("Paragraph 4: newly added.");

        // Perform comparison. The original document will contain revisions.
        original.Compare(revised, "ComparisonAuthor", DateTime.Now);

        // Save the compared document (optional artifact).
        original.Save("Compared.docx");

        // Prepare a summary report of paragraph revisions.
        int insertedCount = 0;
        int deletedCount = 0;
        int formatChangedCount = 0;

        List<string> reportLines = new List<string>
        {
            "Revision Summary Report",
            $"Generated on: {DateTime.Now}",
            ""
        };

        foreach (Revision rev in original.Revisions)
        {
            // Consider only paragraph-level revisions.
            if (rev.ParentNode is not Paragraph paragraph)
                continue;

            string revText = paragraph.GetText()?.Trim() ?? string.Empty;

            switch (rev.RevisionType)
            {
                case RevisionType.Insertion:
                    insertedCount++;
                    reportLines.Add($"[Inserted] \"{revText}\"");
                    break;
                case RevisionType.Deletion:
                    deletedCount++;
                    reportLines.Add($"[Deleted] \"{revText}\"");
                    break;
                case RevisionType.FormatChange:
                    formatChangedCount++;
                    reportLines.Add($"[FormatChanged] \"{revText}\"");
                    break;
                default:
                    // Other revision types are not part of this summary.
                    break;
            }
        }

        reportLines.Add("");
        reportLines.Add($"Inserted paragraphs: {insertedCount}");
        reportLines.Add($"Deleted paragraphs: {deletedCount}");
        reportLines.Add($"Format‑changed paragraphs: {formatChangedCount}");

        // Write the report to a text file.
        string reportPath = "RevisionSummary.txt";
        File.WriteAllLines(reportPath, reportLines);

        // Output the report to the console.
        foreach (string line in reportLines)
        {
            Console.WriteLine(line);
        }
    }
}
