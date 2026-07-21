using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Clone the original to create an edited version and modify its content.
        Document editedDoc = (Document)originalDoc.Clone(true);
        DocumentBuilder editedBuilder = new DocumentBuilder(editedDoc);

        // Modify the second paragraph.
        Paragraph secondParagraph = editedDoc.FirstSection.Body.Paragraphs[1];
        editedBuilder.MoveTo(secondParagraph);
        editedBuilder.Write(" (edited)");

        // Change the text of the third paragraph.
        editedDoc.FirstSection.Body.Paragraphs[2].Runs[0].Text = "Third paragraph - updated.";

        // Compare the original document with the edited version.
        // This adds revisions to the original document representing the differences.
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

        // Build an in‑memory revision report.
        StringBuilder report = new StringBuilder();
        report.AppendLine("Revision Report:");
        report.AppendLine("----------------");

        foreach (Revision rev in originalDoc.Revisions)
        {
            report.AppendLine($"Type   : {rev.RevisionType}");
            report.AppendLine($"Author : {rev.Author}");
            report.AppendLine($"Date   : {rev.DateTime}");

            // For node‑based revisions, include the affected text.
            if (rev.RevisionType != RevisionType.StyleDefinitionChange && rev.ParentNode != null)
                report.AppendLine($"Text   : {rev.ParentNode.GetText().Trim()}");

            report.AppendLine();
        }

        // Output the report to the console.
        Console.Write(report.ToString());
    }
}
