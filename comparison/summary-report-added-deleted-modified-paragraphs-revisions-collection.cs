using System;
using System.IO;
using Aspose.Words;

namespace RevisionSummary
{
    class Program
    {
        static void Main()
        {
            // Load the source document (replace with your actual file path)
            Document doc = new Document("InputWithRevisions.docx");

            // Counters for each revision type that affect paragraphs
            int insertedParagraphs = 0;
            int deletedParagraphs = 0;
            int formatChangedParagraphs = 0;

            // Iterate through all revisions in the document
            foreach (Revision rev in doc.Revisions)
            {
                // We are only interested in revisions whose parent node is a Paragraph
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

            // Build the summary report
            string report = $"Revision Summary Report:{Environment.NewLine}" +
                            $"Inserted paragraphs : {insertedParagraphs}{Environment.NewLine}" +
                            $"Deleted paragraphs  : {deletedParagraphs}{Environment.NewLine}" +
                            $"Modified paragraphs : {formatChangedParagraphs}{Environment.NewLine}";

            // Output the report to the console
            Console.WriteLine(report);

            // Optionally, save the report to a text file
            File.WriteAllText("RevisionSummaryReport.txt", report);

            // Save the (unchanged) document if needed
            doc.Save("OutputDocument.docx");
        }
    }
}
