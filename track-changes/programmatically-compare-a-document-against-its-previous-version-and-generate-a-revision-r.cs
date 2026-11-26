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
        builder.Writeln("Hello World");
        builder.Writeln("This is a sample document.");

        // Save the original document to a memory stream.
        using (MemoryStream originalStream = new MemoryStream())
        {
            originalDoc.Save(originalStream, SaveFormat.Docx);
            originalStream.Position = 0;

            // Load two separate copies: one will stay unchanged, the other will be modified.
            Document unchangedDoc = new Document(originalStream);
            originalStream.Position = 0;
            Document modifiedDoc = new Document(originalStream);

            // Enable track changes and make modifications.
            modifiedDoc.StartTrackRevisions("User1", DateTime.Now);
            DocumentBuilder modBuilder = new DocumentBuilder(modifiedDoc);
            modBuilder.Writeln("Added a new paragraph.");

            // Delete the first paragraph to generate a deletion revision.
            Paragraph firstParagraph = (Paragraph)modifiedDoc.FirstSection.Body.FirstParagraph;
            firstParagraph.Remove();

            modifiedDoc.StopTrackRevisions();

            // Compare the unchanged document against the modified one.
            // The comparison result (with revisions) is stored in unchangedDoc.
            unchangedDoc.Compare(modifiedDoc, "Comparer", DateTime.Now);

            // Build a revision report in memory.
            StringBuilder report = new StringBuilder();
            report.AppendLine("Revision Report:");
            foreach (Revision rev in unchangedDoc.Revisions)
            {
                report.AppendLine($"Type: {rev.RevisionType}");
                report.AppendLine($"Author: {rev.Author}");
                report.AppendLine($"Date: {rev.DateTime}");
                string text = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;
                report.AppendLine($"Text: {text}");
                report.AppendLine("---");
            }

            // Output the report.
            Console.WriteLine(report.ToString());
        }
    }
}
