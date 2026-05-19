using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input Word documents and output PDFs.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample Word documents with a bookmark that encloses the content to extract.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Intro paragraph.");

            // Start bookmark named "Extract".
            builder.StartBookmark("Extract");
            builder.Writeln($"This is the first paragraph inside the extract region of document {i}.");
            builder.Writeln($"This is the second paragraph inside the extract region of document {i}.");
            // End bookmark.
            builder.EndBookmark("Extract");

            builder.Writeln($"Document {i} - Closing paragraph.");

            string docPath = Path.Combine(inputFolder, $"SampleDoc{i}.docx");
            doc.Save(docPath);
        }

        // Process each Word file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(filePath);

            // Locate the bookmark that defines the extraction range.
            Bookmark extractBookmark = sourceDoc.Range.Bookmarks["Extract"];
            if (extractBookmark == null)
                throw new InvalidOperationException($"Bookmark 'Extract' not found in document '{Path.GetFileName(filePath)}'.");

            // Retrieve the text inside the bookmark.
            string extractedText = extractBookmark.Text;

            // Create a new document to hold the extracted content.
            Document resultDoc = new Document();
            DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);
            resultBuilder.Writeln(extractedText.Trim());

            // Determine output PDF file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            string pdfPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}_Extracted.pdf");

            // Save the extracted content as PDF.
            resultDoc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for document '{fileNameWithoutExt}'.");
        }
    }
}
