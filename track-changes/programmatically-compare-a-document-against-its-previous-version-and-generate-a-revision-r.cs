using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This is the second paragraph.");
        builder.Writeln("This is the third paragraph.");

        // Clone the original to create an edited version.
        Document editedDoc = (Document)originalDoc.Clone(true);
        DocumentBuilder editBuilder = new DocumentBuilder(editedDoc);

        // Perform some edits: modify, insert, and delete content.
        // Modify text in the first paragraph.
        Paragraph firstPara = editedDoc.FirstSection.Body.Paragraphs[0];
        firstPara.Runs[0].Text = "This is the UPDATED first paragraph.";

        // Insert a new paragraph.
        editBuilder.MoveToDocumentEnd();
        editBuilder.Writeln("This is an inserted fourth paragraph.");

        // Delete the second paragraph.
        Paragraph secondPara = editedDoc.FirstSection.Body.Paragraphs[1];
        secondPara.Remove();

        // Ensure both documents have no revisions before comparison.
        if (originalDoc.HasRevisions || editedDoc.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the original document with the edited version.
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

        // After comparison, the original document now contains revisions.
        if (!originalDoc.HasRevisions)
            throw new InvalidOperationException("Comparison did not produce any revisions.");

        // Build a revision report in memory.
        StringBuilder reportBuilder = new StringBuilder();
        reportBuilder.AppendLine("Revision Report:");
        reportBuilder.AppendLine("----------------");

        foreach (Revision revision in originalDoc.Revisions)
        {
            // Gather revision details.
            string author = revision.Author;
            DateTime date = revision.DateTime;
            RevisionType type = revision.RevisionType;
            string changedText = revision.ParentNode?.GetText().Trim() ?? string.Empty;

            reportBuilder.AppendLine($"Author: {author}");
            reportBuilder.AppendLine($"Date: {date}");
            reportBuilder.AppendLine($"Type: {type}");
            reportBuilder.AppendLine($"Changed Text: \"{changedText}\"");
            reportBuilder.AppendLine();
        }

        // The report is now stored in the string variable.
        string revisionReport = reportBuilder.ToString();

        // Output the report to the console (optional, demonstrates the result).
        Console.WriteLine(revisionReport);

        // Optionally, accept all revisions to transform the original into the edited version.
        originalDoc.AcceptAllRevisions();

        // Save the final document to a memory stream (demonstrates saving without file I/O).
        using (MemoryStream stream = new MemoryStream())
        {
            originalDoc.Save(stream, SaveFormat.Docx);
            // The stream now contains the final document bytes.
        }
    }
}
