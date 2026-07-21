using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchAppendExample
{
    public static void Main()
    {
        // Base directory for the sample data.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "SampleData");

        // Ensure a clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(baseDir);

        // Create two subfolders each containing a few DOCX files.
        CreateSampleFolder(Path.Combine(baseDir, "FolderA"), "A", 2);
        CreateSampleFolder(Path.Combine(baseDir, "FolderB"), "B", 3);

        // Destination (master) document. It starts with a single empty section.
        Document masterDoc = new Document();

        // Find all DOCX files in all subfolders.
        string[] docxFiles = Directory.GetFiles(baseDir, "*.docx", SearchOption.AllDirectories);

        // Append each source document using UseDestinationStyles.
        foreach (string filePath in docxFiles)
        {
            Document srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        }

        // Export the merged document to PDF.
        string outputPdf = Path.Combine(baseDir, "MergedOutput.pdf");
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Simple validation.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Verify that the merged document contains sections from all source files.
        // masterDoc already contains one initial empty section, so we add 1 to the expected count.
        int expectedSections = docxFiles.Sum(f => new Document(f).Sections.Count) + 1;
        if (masterDoc.Sections.Count != expectedSections)
            throw new InvalidOperationException("The merged document does not contain the expected number of sections.");

        // Indicate successful completion (no interactive console I/O required).
        Console.WriteLine("Batch append completed successfully. PDF saved at:");
        Console.WriteLine(outputPdf);
    }

    // Helper method to create a subfolder with a given number of sample DOCX files.
    private static void CreateSampleFolder(string folderPath, string prefix, int fileCount)
    {
        Directory.CreateDirectory(folderPath);
        for (int i = 1; i <= fileCount; i++)
        {
            string fileName = $"{prefix}_Doc{i}.docx";
            string filePath = Path.Combine(folderPath, fileName);
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i} in folder {Path.GetFileName(folderPath)}.");
            doc.Save(filePath, SaveFormat.Docx);
        }
    }
}
