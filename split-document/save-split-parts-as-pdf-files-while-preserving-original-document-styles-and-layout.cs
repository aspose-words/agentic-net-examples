using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentToPdf
{
    public static void Main()
    {
        // Define a folder for all files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample DOCX document with multiple sections.
        string sourcePath = Path.Combine(artifactsDir, "Sample.docx");
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Split the document by sections and save each part as a PDF.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new blank document.
            Document partDoc = new Document();

            // Import the current section into the new document, preserving formatting.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(importedSection);

            // Define the output PDF file name.
            string pdfPath = Path.Combine(artifactsDir, $"Part_{i + 1}.pdf");

            // Save the part as PDF.
            partDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Validate that all expected PDF files were created.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string pdfPath = Path.Combine(artifactsDir, $"Part_{i + 1}.pdf");
            if (!File.Exists(pdfPath))
                throw new Exception($"Expected split PDF not found: {pdfPath}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Document split into PDF parts successfully.");
    }

    // Helper method to create a sample document with three sections.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - Introduction");
        builder.Writeln("This is the first section of the sample document.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Body");
        builder.Writeln("Content of the second section goes here.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - Conclusion");
        builder.Writeln("Final remarks in the third section.");

        // Save the sample document.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
