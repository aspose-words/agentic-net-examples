using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Use a folder relative to the executable. It will be created if it does not exist.
        string inputFolder = Path.Combine(AppContext.BaseDirectory, "Input");
        Directory.CreateDirectory(inputFolder);

        // Get all .docx files in the folder (adjust the pattern if other formats are needed)
        string[] files = Directory.GetFiles(inputFolder, "*.docx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Load the document from file
            Document doc = new Document(filePath);

            // If the document has any watermark, remove it
            if (doc.Watermark.Type != WatermarkType.None)
                doc.Watermark.Remove();

            // Save the modified document back to the same file
            doc.Save(filePath);
        }

        Console.WriteLine($"Processed {files.Length} file(s) in \"{inputFolder}\".");
    }
}
