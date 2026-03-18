using System;
using System.IO;
using System.Text;
using Aspose.Words;

namespace RevisionReportExample
{
    public class RevisionReportGenerator
    {
        // Generates a revision report comparing two document versions loaded from memory streams.
        // Returns the report as a string; the compared document remains in memory.
        public string GenerateReport(byte[] originalBytes, byte[] editedBytes)
        {
            // Load the original and edited documents from the provided byte arrays.
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            using (MemoryStream editedStream = new MemoryStream(editedBytes))
            {
                Document originalDoc = new Document(originalStream);
                Document editedDoc = new Document(editedStream);

                // Ensure both documents have no existing revisions before comparison.
                if (originalDoc.Revisions.Count == 0 && editedDoc.Revisions.Count == 0)
                {
                    // Compare the documents; revisions will be added to originalDoc.
                    originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);
                }

                // Build the revision report in memory.
                StringBuilder reportBuilder = new StringBuilder();

                foreach (Revision rev in originalDoc.Revisions)
                {
                    reportBuilder.AppendLine($"Revision Type: {rev.RevisionType}");
                    reportBuilder.AppendLine($"Author: {rev.Author}");
                    reportBuilder.AppendLine($"Date: {rev.DateTime}");
                    // Get the text of the node that the revision applies to.
                    string changedText = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;
                    reportBuilder.AppendLine($"Changed Text: \"{changedText}\"");
                    reportBuilder.AppendLine(); // Blank line between revisions
                }

                // Return the complete report.
                return reportBuilder.ToString();
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – replace the paths with your own .docx files.
            string originalPath = "original.docx";
            string editedPath = "edited.docx";

            if (!File.Exists(originalPath) || !File.Exists(editedPath))
            {
                Console.WriteLine("Please ensure both original.docx and edited.docx exist in the executable directory.");
                return;
            }

            byte[] originalBytes = File.ReadAllBytes(originalPath);
            byte[] editedBytes = File.ReadAllBytes(editedPath);

            var generator = new RevisionReportGenerator();
            string report = generator.GenerateReport(originalBytes, editedBytes);

            Console.WriteLine("--- Revision Report ---");
            Console.WriteLine(report);
        }
    }
}
