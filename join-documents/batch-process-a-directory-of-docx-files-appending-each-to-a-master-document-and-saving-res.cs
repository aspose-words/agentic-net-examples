using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define paths for the input folder and the merged PDF output.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");

        // Ensure a clean environment: recreate the input folder.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        Directory.CreateDirectory(inputFolder);

        // Create sample DOCX files that will be merged.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}");

            // Add a page break after each sample except the last.
            if (i < sampleCount)
                builder.InsertBreak(BreakType.PageBreak);

            string filePath = Path.Combine(inputFolder, $"Sample{i}.docx");
            doc.Save(filePath, SaveFormat.Docx);
        }

        // Get all DOCX files from the folder.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        if (docxFiles.Length == 0)
            throw new InvalidOperationException("No source documents were found to merge.");

        // Load the first document as the master document.
        Document masterDoc = new Document(docxFiles[0]);

        // Append the remaining documents to the master document.
        for (int i = 1; i < docxFiles.Length; i++)
        {
            Document srcDoc = new Document(docxFiles[i]);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document as PDF.
        masterDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("Merged PDF was not created.");

        // Validation: ensure the merged document contains a section for each source document.
        if (masterDoc.Sections.Count != docxFiles.Length)
            throw new InvalidOperationException("Merged document does not contain all source sections.");

        // Output a success message (non‑interactive).
        Console.WriteLine($"Successfully merged {docxFiles.Length} documents into '{outputPdfPath}'.");
    }
}
