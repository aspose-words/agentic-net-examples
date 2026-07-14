using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentToPdf
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "SplitPdfParts");
        Directory.CreateDirectory(outputFolder);

        // Create a sample source document with multiple sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - Introduction");
        builder.Writeln("This is the first section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Details");
        builder.Writeln("Details go here.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - Conclusion");
        builder.Writeln("Final remarks.");

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(outputFolder, "SourceDocument.docx");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Split the document by sections and save each part as a PDF.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document partDoc = new Document();

            // Import the current section from the source document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true, ImportFormatMode.KeepSourceFormatting);
            partDoc.Sections.Add(importedSection);

            // Define the output PDF file name.
            string pdfPath = Path.Combine(outputFolder, $"Part_{i + 1}.pdf");

            // Save the split part as PDF, preserving styles and layout.
            partDoc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF part: {pdfPath}");
        }

        // All split PDF files are now available in the output folder.
        Console.WriteLine($"Document split into {sourceDoc.Sections.Count} PDF parts at: {outputFolder}");
    }
}
