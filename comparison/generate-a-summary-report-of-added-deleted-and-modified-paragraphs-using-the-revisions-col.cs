using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");

        // Save the original document (optional, just for reference).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Clone the original to create the revised version.
        Document revised = (Document)original.Clone(true);

        // Modify the text of the second paragraph.
        Paragraph para2 = revised.FirstSection.Body.Paragraphs[1];
        if (para2.Runs.Count > 0)
            para2.Runs[0].Text = "Paragraph 2 - modified";

        // Apply a format change to the first paragraph (e.g., change style to Heading1).
        Paragraph para1 = revised.FirstSection.Body.Paragraphs[0];
        para1.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Delete the third paragraph.
        Paragraph para3 = revised.FirstSection.Body.Paragraphs[2];
        para3.Remove();

        // Add a new fourth paragraph.
        DocumentBuilder revBuilder = new DocumentBuilder(revised);
        revBuilder.Writeln("Paragraph 4");

        // Save the revised document (optional).
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "Revised.docx");
        revised.Save(revisedPath);

        // Compare the documents. Revisions will be stored in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Count added, deleted, and modified (format change) paragraph revisions.
        int added = 0;
        int deleted = 0;
        int modified = 0;

        foreach (Revision rev in original.Revisions)
        {
            // Ensure the revision is related to a paragraph node.
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
                    case RevisionType.FormatChange:
                        modified++;
                        break;
                }
            }
        }

        // Build a simple text report.
        string report = $"Revision Summary:{Environment.NewLine}" +
                        $"Added paragraphs: {added}{Environment.NewLine}" +
                        $"Deleted paragraphs: {deleted}{Environment.NewLine}" +
                        $"Modified paragraphs (format changes): {modified}{Environment.NewLine}";

        // Output the report to the console.
        Console.WriteLine(report);

        // Save the report to a local text file.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionSummary.txt");
        File.WriteAllText(reportPath, report);

        // Save the compared document that contains the revision markup.
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(comparedPath);
    }
}
