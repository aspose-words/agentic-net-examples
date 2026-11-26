using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = "SplitOutput";
        Directory.CreateDirectory(outputDir);

        // Create a sample source document with two sections, each having its own header/footer
        string sourcePath = "SampleDocument.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header - Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer - Section 1");
        builder.MoveToDocumentStart();
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");

        // Start Section 2
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header - Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer - Section 2");

        // Save the source document
        sourceDoc.Save(sourcePath);

        // Load the document (could reuse sourceDoc, but loading demonstrates typical workflow)
        Document doc = new Document(sourcePath);

        // Split the document by sections and save each part as a PDF
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];

            // Create a new empty document for the split part
            Document splitDoc = new Document();

            // Import the section into the new document, preserving formatting
            NodeImporter importer = new NodeImporter(doc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);

            // Append the imported section to the new document
            splitDoc.AppendChild(importedSection);

            // Save the split part as PDF
            string pdfPath = Path.Combine(outputDir, $"Section_{i + 1}.pdf");
            splitDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Validate that the expected PDF files were created
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"Section_{i + 1}.pdf");
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected split PDF not found: {expectedPath}");
        }

        // Optional: indicate success (no interactive input required)
        Console.WriteLine("Document split completed successfully.");
    }
}
