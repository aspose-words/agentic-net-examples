using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentToPdf
{
    public static void Main()
    {
        // Prepare directories.
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "SplitOutput");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three sections, each having its own header/footer.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 3; i++)
        {
            // Set a distinct header for the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write($"Header of Section {i}");

            // Move back to the main body and add content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"This is the content of section {i}.");
            builder.Writeln($"More text in section {i} to demonstrate layout preservation.");

            // Insert a section break after each section except the last one.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Optional: save the original document for reference.
        string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Split the document by sections and save each part as a PDF.
        for (int index = 0; index < sourceDoc.Sections.Count; index++)
        {
            // Create a new empty document.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the current section from the source document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[index], true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(importedSection);

            // Define the output PDF file name.
            string pdfPath = Path.Combine(outputDir, $"Section_{index + 1}.pdf");

            // Save the split part as PDF, preserving styles and layout.
            partDoc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for section {index + 1}.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split into PDF parts successfully.");
    }
}
