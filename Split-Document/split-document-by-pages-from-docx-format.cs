using System;
using System.IO;
using Aspose.Words;

class SplitByPages
{
    static void Main(string[] args)
    {
        // Simple argument handling: source file and output folder.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: SplitByPages <source-docx-path> <output-folder>");
            return;
        }

        string sourcePath = args[0];
        string outputFolder = args[1];

        try
        {
            SplitDocument(sourcePath, outputFolder);
            Console.WriteLine($"Document split into pages successfully. Files are saved in \"{outputFolder}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void SplitDocument(string sourcePath, string outputFolder)
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document(sourcePath);

        // Create the output directory if it does not exist.
        Directory.CreateDirectory(outputFolder);

        // Get the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Iterate through each page, extract it, and save as a separate DOCX file.
        for (int i = 0; i < totalPages; i++)
        {
            // Extract a single page (zero‑based index).
            Document pageDoc = sourceDoc.ExtractPages(i, 1);

            // Construct the output file name (e.g., Page_1.docx, Page_2.docx, ...).
            string outFile = Path.Combine(outputFolder, $"Page_{i + 1}.docx");

            // Save the extracted page.
            pageDoc.Save(outFile);
        }
    }
}
