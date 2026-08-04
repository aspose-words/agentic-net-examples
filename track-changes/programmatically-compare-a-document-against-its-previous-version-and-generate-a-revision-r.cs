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
        builder.Writeln("Second line of the original.");

        // Clone the original to create an edited version.
        Document edited = (Document)original.Clone(true);
        DocumentBuilder editedBuilder = new DocumentBuilder(edited);

        // Modify the edited document: change existing text and add a new line.
        edited.FirstSection.Body.Paragraphs[0].Runs[0].Text = "This is the edited document.";
        editedBuilder.Writeln("An additional line added in the edited version.");

        // Ensure both documents have no revisions before comparison.
        if (original.HasRevisions || edited.HasRevisions)
            throw new InvalidOperationException("Documents should not contain revisions before comparison.");

        // Compare the original document with the edited version.
        original.Compare(edited, "Comparer", DateTime.Now);

        // Build a revision report in memory.
        StringBuilder report = new StringBuilder();
        report.AppendLine("Revision Report:");
        report.AppendLine("----------------");

        foreach (Revision revision in original.Revisions)
        {
            report.AppendLine($"Author: {revision.Author}");
            report.AppendLine($"Date: {revision.DateTime}");
            report.AppendLine($"Type: {revision.RevisionType}");
            report.AppendLine($"Changed Text: \"{revision.ParentNode.GetText().Trim()}\"");
            report.AppendLine();
        }

        // Output the report to the console.
        Console.WriteLine(report.ToString());
    }
}
