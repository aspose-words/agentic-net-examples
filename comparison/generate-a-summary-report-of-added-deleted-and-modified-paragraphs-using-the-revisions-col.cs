using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for sample documents and results.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        string editedPath = Path.Combine(outputDir, "Edited.docx");
        string comparedPath = Path.Combine(outputDir, "Compared.docx");
        string reportPath = Path.Combine(outputDir, "RevisionSummary.txt");

        // ---------- Create original document ----------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        original.Save(originalPath);

        // ---------- Create edited document with differences ----------
        Document edited = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(edited);
        builderEdited.Writeln("Paragraph 1");                 // unchanged
        // Paragraph 2 is removed – no corresponding line.
        builderEdited.Writeln("Paragraph 3 modified");        // changed text
        builderEdited.Writeln("Paragraph 4");                // new paragraph
        edited.Save(editedPath);

        // ---------- Compare documents ----------
        // The original document will receive revisions describing the differences.
        original.Compare(edited, "Comparer", DateTime.Now);
        original.Save(comparedPath);

        // ---------- Analyze revisions ----------
        int added = 0;
        int deleted = 0;

        foreach (Revision rev in original.Revisions)
        {
            // Consider only paragraph revisions.
            if (rev.ParentNode?.NodeType == NodeType.Paragraph)
            {
                switch (rev.RevisionType)
                {
                    case RevisionType.Insertion:
                        added++;
                        break;
                    case RevisionType.Deletion:
                        deleted++;
                        break;
                }
            }
        }

        // Approximate modified paragraphs as pairs of deletion+insertion.
        int modified = Math.Min(added, deleted);

        // ---------- Write summary report ----------
        StringBuilder report = new StringBuilder();
        report.AppendLine($"Added paragraphs: {added}");
        report.AppendLine($"Deleted paragraphs: {deleted}");
        report.AppendLine($"Modified paragraphs: {modified}");

        File.WriteAllText(reportPath, report.ToString());

        // Optional console output for quick verification.
        Console.WriteLine(report.ToString());
    }
}
