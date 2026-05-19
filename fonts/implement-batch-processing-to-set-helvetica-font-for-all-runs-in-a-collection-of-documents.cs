using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure both folders exist to avoid DirectoryNotFoundException.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Retrieve all .docx files from the input folder.
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx", SearchOption.TopDirectoryOnly);

        // If there are no documents, inform the user and exit gracefully.
        if (docFiles.Length == 0)
        {
            Console.WriteLine($"No .docx files found in '{inputFolder}'.");
            Console.WriteLine("Batch processing completed.");
            return;
        }

        // Process each document.
        foreach (string filePath in docFiles)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Get all Run nodes in the document.
            Aspose.Words.NodeCollection runs = doc.GetChildNodes(Aspose.Words.NodeType.Run, true);

            // Set the font of each Run to Helvetica and validate.
            foreach (Aspose.Words.Run run in runs)
            {
                run.Font.Name = "Helvetica";

                if (!string.Equals(run.Font.Name, "Helvetica", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"Failed to set font for a run in document '{filePath}'.");
                }
            }

            // Determine the output file path and save the modified document.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
            {
                throw new FileNotFoundException($"The output file was not created: {outputPath}");
            }

            Console.WriteLine($"Processed '{Path.GetFileName(filePath)}' and saved to '{outputPath}'.");
        }

        Console.WriteLine("Batch processing completed.");
    }
}
