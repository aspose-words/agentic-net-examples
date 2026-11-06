using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentToPdf
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with multiple sections, each having its own header/footer.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Split the document by sections and save each part as a PDF.
        int partIndex = 1;
        foreach (Section section in sourceDoc.Sections)
        {
            // Create a new empty document and clear its default content.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the automatically created empty section.

            // Import the current section into the new document, preserving formatting.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            partDoc.AppendChild(importedSection);

            // Save the split part as PDF.
            string pdfPath = Path.Combine(outputDir, $"Part_{partIndex}.pdf");
            partDoc.Save(pdfPath);

            // Verify the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");

            partIndex++;
        }

        // Ensure at least one PDF was generated.
        if (partIndex == 1)
            throw new InvalidOperationException("No sections were found to split.");
    }

    // Generates a sample DOCX with three sections, each having distinct headers, footers, and body content.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            // Header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln($"Header for Section {i}");

            // Footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln($"Footer for Section {i}");

            // Body content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"This is the content of Section {i}.");
            builder.Writeln($"More text in Section {i} to ensure layout is preserved.");

            // Insert a section break after each section except the last.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the sample document.
        doc.Save(filePath);
    }
}
