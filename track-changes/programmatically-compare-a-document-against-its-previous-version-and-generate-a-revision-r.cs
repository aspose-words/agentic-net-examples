using System;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");
        builder.Writeln("It has two paragraphs.");

        // Create an edited version of the document.
        Document edited = (Document)original.Clone(true);
        DocumentBuilder editBuilder = new DocumentBuilder(edited);

        // Modify the first paragraph.
        Paragraph firstParagraph = edited.FirstSection.Body.FirstParagraph;
        if (firstParagraph.Runs.Count > 0)
            firstParagraph.Runs[0].Text = "This is the edited document.";

        // Add an additional paragraph.
        editBuilder.Writeln("Additional paragraph added.");

        // Compare the edited document to the original.
        // This will generate revisions in the original document.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Build a revision report in memory.
        StringBuilder report = new StringBuilder();
        report.AppendLine($"Total revisions: {original.Revisions.Count}");
        foreach (Revision rev in original.Revisions)
        {
            report.AppendLine($"Author: {rev.Author}");
            report.AppendLine($"Date: {rev.DateTime}");
            report.AppendLine($"Type: {rev.RevisionType}");
            string text = rev.ParentNode != null ? rev.ParentNode.GetText().Trim() : "<no text>";
            report.AppendLine($"Text: {text}");
            report.AppendLine("---");
        }

        // Output the report.
        Console.Write(report.ToString());
    }
}
