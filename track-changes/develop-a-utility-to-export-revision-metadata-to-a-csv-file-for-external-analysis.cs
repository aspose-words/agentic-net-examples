using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This is the second paragraph.");
        builder.Writeln("This is the third paragraph.");

        // Enable track changes.
        string author = "John Doe";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Perform some modifications to generate revisions.
        // Insert a new paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is an inserted paragraph.");

        // Delete the second paragraph.
        Paragraph secondParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        secondParagraph.Remove();

        // Change formatting of the first paragraph (make it bold).
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        firstParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates that the document contains revisions).
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Export revision metadata to CSV.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("RevisionIndex,RevisionType,Author,DateTime,Text");

        RevisionCollection revisions = doc.Revisions;
        for (int i = 0; i < revisions.Count; i++)
        {
            Revision rev = revisions[i];
            string revType = rev.RevisionType.ToString();
            string revAuthor = rev.Author;
            string revDate = rev.DateTime.ToString("o"); // ISO 8601 format
            string revText = rev.ParentNode?.GetText()?.Replace("\r", " ").Replace("\n", " ").Trim();

            // Escape commas in text.
            if (revText != null && revText.Contains(","))
                revText = $"\"{revText}\"";

            csvBuilder.AppendLine($"{i},{revType},{revAuthor},{revDate},{revText}");
        }

        string csvPath = "Revisions.csv";
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
    }
}
