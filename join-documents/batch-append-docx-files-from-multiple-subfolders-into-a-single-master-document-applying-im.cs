using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Base directory for input documents
        string baseInputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        // Ensure a clean environment
        if (Directory.Exists(baseInputDir))
            Directory.Delete(baseInputDir, true);
        Directory.CreateDirectory(baseInputDir);

        // Define subfolders
        string[] subFolders = { "FolderA", "FolderB" };
        int docsPerFolder = 2; // number of DOCX files per subfolder

        // Create sample DOCX files in each subfolder
        foreach (string folder in subFolders)
        {
            string folderPath = Path.Combine(baseInputDir, folder);
            Directory.CreateDirectory(folderPath);

            for (int i = 1; i <= docsPerFolder; i++)
            {
                string fileName = $"Sample_{folder}_{i}.docx";
                string filePath = Path.Combine(folderPath, fileName);

                // Create a simple document with identifiable content
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"This is {fileName} located in {folder}.");
                doc.Save(filePath, SaveFormat.Docx);
            }
        }

        // Prepare the master document (initially empty)
        Document masterDoc = new Document();

        // Keep track of how many source documents are processed
        int sourceDocCount = 0;

        // Batch append all DOCX files from subfolders
        foreach (string folder in subFolders)
        {
            string folderPath = Path.Combine(baseInputDir, folder);
            string[] docxFiles = Directory.GetFiles(folderPath, "*.docx");

            foreach (string docxPath in docxFiles)
            {
                Document srcDoc = new Document(docxPath);
                masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
                sourceDocCount++;
            }
        }

        // Define output paths
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string masterDocPath = Path.Combine(outputDir, "Master.docx");
        string masterPdfPath = Path.Combine(outputDir, "Master.pdf");

        // Save the merged document in DOCX and PDF formats
        masterDoc.Save(masterDocPath, SaveFormat.Docx);
        masterDoc.Save(masterPdfPath, SaveFormat.Pdf);

        // Validation: ensure output files exist
        if (!File.Exists(masterDocPath))
            throw new InvalidOperationException("Merged DOCX file was not created.");
        if (!File.Exists(masterPdfPath))
            throw new InvalidOperationException("Merged PDF file was not created.");

        // Validation: ensure the master document contains sections from all source docs
        // The initial empty document has one section; each appended document adds its sections.
        int expectedSections = 1 + sourceDocCount; // assuming each source has one section
        if (masterDoc.Sections.Count != expectedSections)
            throw new InvalidOperationException($"Expected {expectedSections} sections, but found {masterDoc.Sections.Count}.");

        // Program completed successfully
    }
}
