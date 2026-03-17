using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class RevisionReport
{
    static void Main()
    {
        // Load the source document that contains tracked changes.
        // Replace the path with the actual location of your document.
        string inputPath = @"C:\Docs\Revisions.docx";
        Document doc = new Document(inputPath);

        // Prepare a simple text report.
        // We'll write the report to a new text file.
        string reportPath = @"C:\Docs\RevisionReport.txt";

        using (System.IO.StreamWriter writer = new System.IO.StreamWriter(reportPath, false))
        {
            writer.WriteLine("Revision Report");
            writer.WriteLine("----------------");
            writer.WriteLine();

            // Iterate through all revisions in the document.
            foreach (Revision rev in doc.Revisions)
            {
                // Get revision type and author.
                RevisionType type = rev.RevisionType;
                string author = rev.Author;

                // Determine the paragraph that contains the revision.
                // Most revisions are attached to a node (e.g., Run). Get its ancestor paragraph.
                Paragraph para = rev.ParentNode?.GetAncestor(NodeType.Paragraph) as Paragraph;

                // If the revision is a style change, ParentNode is null; handle gracefully.
                int paragraphNumber = para != null
                    ? doc.FirstSection.Body.Paragraphs.IndexOf(para) + 1 // 1‑based index for readability
                    : -1; // Indicates not applicable

                // Write the information to the report.
                writer.WriteLine($"Type: {type}, Author: {author}, Paragraph #: {(paragraphNumber > 0 ? paragraphNumber.ToString() : "N/A")}");
            }
        }

        // Optionally, display the location of the generated report.
        Console.WriteLine($"Revision report generated at: {reportPath}");
    }
}
