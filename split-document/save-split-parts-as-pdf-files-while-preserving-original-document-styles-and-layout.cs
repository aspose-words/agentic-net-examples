using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentToPdf
{
    public static void Main()
    {
        // Create a sample document with multiple sections, each having its own header/footer.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 3; i++)
        {
            // Start a new section for each iteration (except the first, which is already created).
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Header
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write($"Header for Section {i}");

            // Footer
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write($"Footer for Section {i}");

            // Body content
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Font.Size = 12;
            builder.Writeln($"This is the body of section {i}.");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        }

        // Ensure the source document has up‑to‑date layout information.
        sourceDoc.UpdatePageLayout();

        // Split the document by sections and save each part as a PDF.
        int sectionCount = sourceDoc.Sections.Count;
        for (int idx = 0; idx < sectionCount; idx++)
        {
            Section section = sourceDoc.Sections[idx];

            // Create a new empty document to hold the single section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the section (including its headers/footers) into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Save the split part as PDF.
            string pdfFileName = $"Section_{idx + 1}.pdf";
            splitDoc.Save(pdfFileName, SaveFormat.Pdf);

            // Validate that the PDF file was created.
            if (!File.Exists(pdfFileName))
                throw new InvalidOperationException($"Failed to create split PDF: {pdfFileName}");
        }

        Console.WriteLine("Document split into PDF sections completed.");
    }
}
