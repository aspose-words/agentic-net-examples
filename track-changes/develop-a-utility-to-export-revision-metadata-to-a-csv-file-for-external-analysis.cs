using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
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
        builder.Writeln("This is the original paragraph.");

        // Start tracking revisions.
        string author = "John Doe";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Insert a new paragraph – this will be recorded as an insertion revision.
        builder.Writeln("This paragraph was inserted while tracking changes.");

        // Delete a word from the original paragraph – this will be recorded as a deletion revision.
        // Move the cursor to the first paragraph.
        builder.MoveToParagraph(0, 0);
        // Replace the word "original" with an empty string.
        doc.Range.Replace(new Regex("original"), "", new FindReplaceOptions(FindReplaceDirection.Forward));

        // Change formatting of a word – this will be recorded as a format‑change revision.
        builder.MoveToDocumentEnd();
        builder.Font.Bold = true;
        builder.Writeln("Bold text added.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document that contains the revisions.
        string docPath = "TrackedDocument.docx";
        doc.Save(docPath);

        // Export revision metadata to CSV.
        string csvPath = "Revisions.csv";
        using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            // Write CSV header.
            writer.WriteLine("Index,Type,Author,DateTime,ParentNodeText");

            int index = 0;
            foreach (Revision rev in doc.Revisions)
            {
                // Get the text of the parent node for context and escape quotes.
                string parentText = rev.ParentNode?.GetText()?.Replace("\"", "\"\"") ?? string.Empty;
                parentText = $"\"{parentText}\"";

                string line = $"{index},{rev.RevisionType},{rev.Author},{rev.DateTime:o},{parentText}";
                writer.WriteLine(line);
                index++;
            }
        }

        // Accept all revisions and save the final document.
        doc.AcceptAllRevisions();
        doc.Save("FinalDocument.docx");
    }
}
