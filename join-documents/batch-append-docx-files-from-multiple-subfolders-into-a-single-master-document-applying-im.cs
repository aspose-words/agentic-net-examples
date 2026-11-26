using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Base directory for input documents.
        string baseInputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        // Ensure the base directory exists.
        Directory.CreateDirectory(baseInputDir);

        // Define subfolders to simulate multiple source locations.
        string[] subfolders = { "FolderA", "FolderB" };
        foreach (string subfolder in subfolders)
        {
            string subfolderPath = Path.Combine(baseInputDir, subfolder);
            Directory.CreateDirectory(subfolderPath);

            // Create two sample DOCX files in each subfolder.
            for (int i = 1; i <= 2; i++)
            {
                string fileName = $"Doc{i}.docx";
                string filePath = Path.Combine(subfolderPath, fileName);

                // Build a simple document with identifiable content.
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"Content from {subfolder} - {fileName}");
                doc.Save(filePath, SaveFormat.Docx);
            }
        }

        // Prepare the master document that will receive all appended content.
        Document masterDoc = new Document();

        // Track the total number of sections from source documents for validation.
        int totalSourceSections = 0;

        // Iterate through each subfolder and append its DOCX files.
        foreach (string subfolder in subfolders)
        {
            string subfolderPath = Path.Combine(baseInputDir, subfolder);
            string[] sourceFiles = Directory.GetFiles(subfolderPath, "*.docx", SearchOption.TopDirectoryOnly);

            foreach (string sourceFile in sourceFiles)
            {
                Document sourceDoc = new Document(sourceFile);
                totalSourceSections += sourceDoc.Sections.Count;

                // Append the source document using destination styles.
                masterDoc.AppendDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);
            }
        }

        // Define output paths.
        string outputDocxPath = Path.Combine(Directory.GetCurrentDirectory(), "Master.docx");
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Master.pdf");

        // Save the merged document in DOCX and PDF formats.
        masterDoc.Save(outputDocxPath, SaveFormat.Docx);
        masterDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure output files exist.
        if (!File.Exists(outputDocxPath) || !File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("Merged output files were not created as expected.");
        }

        // Validation: ensure the master document contains the expected number of sections.
        // The master document starts with one empty section; each source document adds its sections.
        int expectedSections = 1 + totalSourceSections;
        if (masterDoc.Sections.Count != expectedSections)
        {
            throw new InvalidOperationException($"Section count mismatch. Expected: {expectedSections}, Actual: {masterDoc.Sections.Count}");
        }
    }
}
