using System;
using System.IO;
using Aspose.Words;

namespace RevisionSummary
{
    class Program
    {
        static void Main()
        {
            // Create a new document with some initial content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Original paragraph 1.");
            builder.Writeln("Original paragraph 2.");
            builder.Writeln("Original paragraph 3.");

            // Enable Track Revisions to capture changes.
            doc.TrackRevisions = true;

            // Insert a new paragraph (Insertion revision).
            builder.MoveToDocumentEnd();
            builder.Writeln("Inserted paragraph.");

            // Delete a paragraph (Deletion revision).
            Paragraph paraToDelete = doc.FirstSection.Body.Paragraphs[0];
            paraToDelete.Remove();

            // Change formatting of a paragraph (FormatChange revision).
            Paragraph paraToFormat = doc.FirstSection.Body.Paragraphs[1];
            paraToFormat.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Disable tracking to stop further revisions.
            doc.TrackRevisions = false;

            // Counters for each revision type that affect paragraphs.
            int insertedParagraphs = 0;
            int deletedParagraphs = 0;
            int formatChangedParagraphs = 0;

            // Iterate through all revisions in the document.
            foreach (Revision rev in doc.Revisions)
            {
                // We are only interested in revisions whose parent node is a Paragraph.
                if (rev.ParentNode?.NodeType == NodeType.Paragraph)
                {
                    switch (rev.RevisionType)
                    {
                        case RevisionType.Insertion:
                            insertedParagraphs++;
                            break;
                        case RevisionType.Deletion:
                            deletedParagraphs++;
                            break;
                        case RevisionType.FormatChange:
                            formatChangedParagraphs++;
                            break;
                    }
                }
            }

            // Build the summary report.
            string report = $"Revision Summary Report:{Environment.NewLine}" +
                            $"Inserted paragraphs : {insertedParagraphs}{Environment.NewLine}" +
                            $"Deleted paragraphs  : {deletedParagraphs}{Environment.NewLine}" +
                            $"Modified paragraphs : {formatChangedParagraphs}{Environment.NewLine}";

            // Output the report to the console.
            Console.WriteLine(report);

            // Optionally, save the report to a text file.
            File.WriteAllText("RevisionSummaryReport.txt", report);

            // Save the document with revisions for inspection (optional).
            doc.Save("OutputDocument.docx");
        }
    }
}
