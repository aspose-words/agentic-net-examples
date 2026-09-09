using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fonts;

namespace FontBatchProcessor
{
    public class Program
    {
        public static void Main()
        {
            // Define input and output directories (relative to the executable location).
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string inputDir = Path.Combine(baseDir, "InputDocs");
            string outputDir = Path.Combine(baseDir, "OutputDocs");

            // Ensure the directories exist.
            Directory.CreateDirectory(inputDir);
            Directory.CreateDirectory(outputDir);

            // Collect all .docx files from the input directory.
            string[] docFiles = Directory.GetFiles(inputDir, "*.docx", SearchOption.TopDirectoryOnly);

            // Process each document.
            foreach (string filePath in docFiles)
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Iterate over all Run nodes and set the font to Helvetica.
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                foreach (Run run in runs)
                {
                    // Set the font name.
                    run.Font.Name = "Helvetica";

                    // Validate that the font name was set correctly.
                    if (!string.Equals(run.Font.Name, "Helvetica", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"Failed to set font for run in document '{filePath}'.");
                    }
                }

                // Save the modified document to the output directory.
                string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
                doc.Save(outputPath);

                // Verify that the file was created.
                if (!File.Exists(outputPath))
                {
                    throw new FileNotFoundException($"The output file was not created: {outputPath}");
                }

                Console.WriteLine($"Processed '{Path.GetFileName(filePath)}' and saved to output folder.");
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
