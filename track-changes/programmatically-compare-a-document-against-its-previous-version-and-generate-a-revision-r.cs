using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document and add some content.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original paragraph.");
        builder.Writeln("Another line in the original document.");

        // Clone the original to create the edited version.
        Document edited = (Document)original.Clone(true);
        // Modify the edited document: change text and add a new paragraph.
        Paragraph firstParagraph = edited.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "This is the edited paragraph.";
        edited.FirstSection.Body.AppendParagraph("An added line in the edited document.");

        // Ensure both documents have no revisions before comparison.
        if (original.HasRevisions || edited.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the original document with the edited version.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Build a revision report in memory.
        StringBuilder report = new StringBuilder();
        report.AppendLine("Revision Report:");
        report.AppendLine($"Total revisions: {original.Revisions.Count}");
        foreach (Revision rev in original.Revisions)
        {
            report.AppendLine($"- Type: {rev.RevisionType}");
            report.AppendLine($"  Author: {rev.Author}");
            report.AppendLine($"  Date: {rev.DateTime}");
            // For revisions that affect nodes, output the affected text.
            if (rev.ParentNode != null)
                report.AppendLine($"  Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Output the report to the console.
        Console.WriteLine(report.ToString());

        // (Optional) Save the original document with revisions to a memory stream.
        using (MemoryStream ms = new MemoryStream())
        {
            original.Save(ms, SaveFormat.Docx);
            // The stream can be used further if needed.
        }
    }
}
