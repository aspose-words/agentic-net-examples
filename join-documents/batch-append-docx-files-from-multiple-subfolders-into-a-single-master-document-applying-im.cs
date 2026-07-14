using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for sample input documents.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        // Ensure a clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(baseDir);

        // Create two subfolders.
        string[] subFolders = { "FolderA", "FolderB" };
        foreach (string folder in subFolders)
        {
            string folderPath = Path.Combine(baseDir, folder);
            Directory.CreateDirectory(folderPath);

            // Create two sample DOCX files in each subfolder.
            for (int i = 1; i <= 2; i++)
            {
                string fileName = $"Doc{i}_{folder}.docx";
                string filePath = Path.Combine(folderPath, fileName);
                CreateSampleDocument(filePath, $"Content of {fileName}");
            }
        }

        // Master document that will receive all appended documents.
        Document masterDoc = new Document();

        // Find all DOCX files in the subfolders.
        string[] docxFiles = Directory.GetFiles(baseDir, "*.docx", SearchOption.AllDirectories);

        // Append each document using UseDestinationStyles.
        foreach (string file in docxFiles)
        {
            Document srcDoc = new Document(file);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        }

        // Save the merged document as PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The merged PDF was not created.");

        // Validation: ensure the master document contains sections from all source docs.
        // A blank master document starts with one section; each appended document adds its sections.
        int expectedSections = 1 + docxFiles.Length; // assuming each source has one section.
        if (masterDoc.Sections.Count != expectedSections)
            throw new InvalidOperationException($"Expected {expectedSections} sections, but found {masterDoc.Sections.Count}.");

        // Optionally, write a simple confirmation to the console (no user interaction required).
        Console.WriteLine($"Successfully merged {docxFiles.Length} documents into '{outputPdf}'.");
    }

    // Helper method to create a simple DOCX file with given text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
