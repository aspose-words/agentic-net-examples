using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Build a sample document with three sections, each having its own
        //    header, footer and body content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 1");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 1.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 2");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 2.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 3");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 3");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 3.");

        // -----------------------------------------------------------------
        // 2. Split the source document by its sections.
        //    For each section create a new document, import the section,
        //    and save it as a PDF file.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section sourceSection = sourceDoc.Sections[i];

            // Create a new empty document and clear its default section.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // removes the default empty section.

            // Import the source section into the new document, preserving formatting.
            Section importedSection = (Section)partDoc.ImportNode(
                sourceSection, true, ImportFormatMode.KeepSourceFormatting);

            // Append the imported section as the sole section of the new document.
            partDoc.AppendChild(importedSection);

            // Define the output PDF file name.
            string pdfPath = Path.Combine(outputDir, $"Section_{i + 1}.pdf");

            // Save the part as PDF.
            partDoc.Save(pdfPath);
        }

        // -----------------------------------------------------------------
        // 3. Verify that the expected PDF files were created.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"Section_{i + 1}.pdf");
            if (!File.Exists(expectedPath))
                throw new InvalidOperationException($"Expected split PDF not found: {expectedPath}");
        }

        // Program completed successfully.
    }
}
