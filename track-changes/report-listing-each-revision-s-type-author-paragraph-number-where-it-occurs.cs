using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

class RevisionReport
{
    static void Main()
    {
        // Create a new document in memory and add some tracked changes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions.
        doc.StartTrackRevisions("Author1");

        // Add some paragraphs – these insertions will be recorded as revisions.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Prepare a simple text report.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "RevisionReport.txt");

        using (StreamWriter writer = new StreamWriter(reportPath, false))
        {
            writer.WriteLine("Revision Report");
            writer.WriteLine("----------------");
            writer.WriteLine();

            // Iterate through all revisions in the document.
            foreach (Revision rev in doc.Revisions)
            {
                RevisionType type = rev.RevisionType;
                string author = rev.Author;

                // Determine the paragraph that contains the revision.
                Paragraph para = rev.ParentNode?.GetAncestor(NodeType.Paragraph) as Paragraph;

                int paragraphNumber = para != null
                    ? doc.FirstSection.Body.Paragraphs.IndexOf(para) + 1 // 1‑based index for readability
                    : -1; // Indicates not applicable

                writer.WriteLine(
                    $"Type: {type}, Author: {author}, Paragraph #: {(paragraphNumber > 0 ? paragraphNumber.ToString() : "N/A")}");
            }
        }

        Console.WriteLine($"Revision report generated at: {reportPath}");
    }
}
