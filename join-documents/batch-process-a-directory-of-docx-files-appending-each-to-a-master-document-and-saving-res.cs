using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for the example files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsExample");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputFile = Path.Combine(baseDir, "MasterDocument.pdf");

        // Clean any previous run and create required folders.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);

        // -----------------------------------------------------------------
        // Create sample source DOCX files inside the input folder.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string srcPath = Path.Combine(inputDir, $"Doc{i}.docx");
            Document srcDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(srcDoc);
            builder.Writeln($"This is the content of document {i}.");
            srcDoc.Save(srcPath, SaveFormat.Docx);
        }

        // -----------------------------------------------------------------
        // Create the master document that will receive all appended files.
        // Ensure it starts with no sections so that the final count matches the
        // number of source documents.
        // -----------------------------------------------------------------
        Document masterDoc = new Document();
        // Remove the default empty section.
        masterDoc.RemoveAllChildren();
        masterDoc.Sections.Clear();

        // Get all DOCX files from the input folder.
        string[] sourceFiles = Directory.GetFiles(inputDir, "*.docx");

        foreach (string srcPath in sourceFiles)
        {
            // Load the source document.
            Document srcDoc = new Document(srcPath);

            // Insert a page break before appending, except before the first document.
            if (masterDoc.Sections.Count > 0)
            {
                DocumentBuilder mb = new DocumentBuilder(masterDoc);
                mb.MoveToDocumentEnd();
                mb.InsertBreak(BreakType.PageBreak);
            }

            // Append the source document while preserving its formatting.
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // -----------------------------------------------------------------
        // Save the merged result as PDF.
        // -----------------------------------------------------------------
        masterDoc.Save(outputFile, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure the PDF was created and sections count matches.
        // -----------------------------------------------------------------
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The merged PDF file was not created.");

        int expectedSections = sourceFiles.Length;
        if (masterDoc.Sections.Count != expectedSections)
            throw new InvalidOperationException($"Expected {expectedSections} sections, but found {masterDoc.Sections.Count}.");

        Console.WriteLine($"Merged PDF successfully created at: {outputFile}");
    }
}
