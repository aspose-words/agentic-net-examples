using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string outputDir = Path.Combine(artifactsDir, "SplitPdfParts");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with multiple sections, each having its
        //    own header and body content. This document will be the source for
        //    the split operation.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 3; i++)
        {
            // Write body content for the current section.
            builder.Writeln($"This is the content of section {i}.");

            // Add a header that belongs to the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln($"Header for section {i}");

            // Return the cursor to the main body.
            builder.MoveToDocumentEnd();

            // Insert a section break after all but the last section.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document (optional, useful for inspection).
        string sourcePath = Path.Combine(artifactsDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the source document by its sections.
        //    For each section we create a new Document, import the section
        //    (including its headers/footers), and save it as a PDF.
        // -----------------------------------------------------------------
        int partNumber = 1;
        foreach (Section section in sourceDoc.Sections)
        {
            // Create an empty document that will hold the single section.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the section from the source document, preserving formatting.
            Section importedSection = (Section)partDoc.ImportNode(section, true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(importedSection);

            // Define the output PDF file name.
            string partPath = Path.Combine(outputDir, $"Part_{partNumber}.pdf");

            // Save the part as PDF.
            partDoc.Save(partPath, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split PDF: {partPath}");

            partNumber++;
        }

        // -----------------------------------------------------------------
        // 3. Simple validation: ensure the expected number of PDF files exist.
        // -----------------------------------------------------------------
        int expectedParts = sourceDoc.Sections.Count;
        int actualParts = Directory.GetFiles(outputDir, "*.pdf").Length;
        if (actualParts != expectedParts)
            throw new InvalidOperationException($"Expected {expectedParts} PDF parts, but found {actualParts}.");

        // Program completed successfully.
    }
}
