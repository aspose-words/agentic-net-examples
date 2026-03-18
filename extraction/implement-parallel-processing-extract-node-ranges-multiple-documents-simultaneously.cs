using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;

namespace AsposeWordsParallelProcessing
{
    public static class DocumentRangeExtractor
    {
        /// <summary>
        /// Extracts the text of each Section's range from a collection of Word documents in parallel.
        /// The extracted text for each document is saved to a new .txt file next to the source file.
        /// </summary>
        /// <param name="sourceFiles">Full paths of the source .doc/.docx files.</param>
        public static void ExtractSectionRangesParallel(IEnumerable<string> sourceFiles)
        {
            // Parallel.ForEach will schedule the work on multiple threads automatically.
            Parallel.ForEach(sourceFiles, sourcePath =>
            {
                // Load the document (lifecycle rule: load).
                Document doc = new Document(sourcePath);

                // Build a string that contains the text of every section's range.
                // Using StringBuilder for efficiency when concatenating many sections.
                var builder = new System.Text.StringBuilder();

                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    // Each Section is a Node; its Range property gives a Range object.
                    // The Range.Text property returns the full text covered by that node.
                    string sectionText = doc.Sections[i].Range.Text?.Trim() ?? string.Empty;

                    builder.AppendLine($"--- Section {i + 1} ---");
                    builder.AppendLine(sectionText);
                    builder.AppendLine(); // Blank line between sections.
                }

                // Determine output file path (same folder, same name with .txt extension).
                string outputPath = Path.ChangeExtension(sourcePath, ".txt");

                // Save the extracted text (lifecycle rule: save).
                File.WriteAllText(outputPath, builder.ToString());

                // Optional: log progress (console output is thread‑safe for WriteLine).
                Console.WriteLine($"Processed '{Path.GetFileName(sourcePath)}' -> '{Path.GetFileName(outputPath)}'");
            });
        }

        // Example usage.
        public static void Main()
        {
            // Example list of document paths. Replace with actual file locations.
            var docs = new List<string>
            {
                @"C:\Docs\Document1.docx",
                @"C:\Docs\Document2.docx",
                @"C:\Docs\Document3.docx"
            };

            ExtractSectionRangesParallel(docs);
        }
    }
}
