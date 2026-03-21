using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
            Parallel.ForEach(sourceFiles, sourcePath =>
            {
                // Load the document.
                Document doc = new Document(sourcePath);

                // Build a string that contains the text of every section's range.
                var builder = new StringBuilder();

                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    string sectionText = doc.Sections[i].Range.Text?.Trim() ?? string.Empty;

                    builder.AppendLine($"--- Section {i + 1} ---");
                    builder.AppendLine(sectionText);
                    builder.AppendLine(); // Blank line between sections.
                }

                // Determine output file path (same folder, same name with .txt extension).
                string outputPath = Path.ChangeExtension(sourcePath, ".txt");

                // Save the extracted text.
                File.WriteAllText(outputPath, builder.ToString());

                // Log progress.
                Console.WriteLine($"Processed '{Path.GetFileName(sourcePath)}' -> '{Path.GetFileName(outputPath)}'");
            });
        }

        // Example usage.
        public static void Main()
        {
            // Create a temporary directory to hold sample documents.
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDocsSample");
            Directory.CreateDirectory(tempDir);

            // Generate sample .docx files.
            var docs = new List<string>();
            for (int i = 1; i <= 3; i++)
            {
                string docPath = Path.Combine(tempDir, $"Document{i}.docx");
                CreateSampleDocument(docPath, i);
                docs.Add(docPath);
            }

            // Run the parallel extraction.
            ExtractSectionRangesParallel(docs);

            // Optionally display the generated text files.
            Console.WriteLine("\n--- Extracted Text Files ---");
            foreach (var docPath in docs)
            {
                string txtPath = Path.ChangeExtension(docPath, ".txt");
                Console.WriteLine($"\n{Path.GetFileName(txtPath)}:");
                Console.WriteLine(File.ReadAllText(txtPath));
            }
        }

        // Helper method to create a simple Word document with a few sections.
        private static void CreateSampleDocument(string path, int docNumber)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Section 1
            builder.Writeln($"Document {docNumber} - Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 2
            builder.Writeln($"Document {docNumber} - Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 3
            builder.Writeln($"Document {docNumber} - Section 3");

            doc.Save(path);
        }
    }
}
