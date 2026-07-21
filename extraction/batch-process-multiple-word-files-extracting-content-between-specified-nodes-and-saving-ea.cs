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
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");

        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Intro paragraph.");
            builder.StartBookmark("Extract");
            builder.Writeln($"This is the extracted content of document {i}.");
            builder.Writeln("It may contain multiple paragraphs.");
            builder.EndBookmark("Extract");
            builder.Writeln($"Document {i} - Closing paragraph.");

            string filePath = Path.Combine(inputDir, $"Sample{i}.docx");
            doc.Save(filePath);
        }

        // Process each document: extract bookmark content and save as PDF.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document sourceDoc = new Document(filePath);

            // Locate the bookmark named "Extract".
            Bookmark bookmark = sourceDoc.Range.Bookmarks["Extract"];
            if (bookmark == null)
                throw new InvalidOperationException($"Bookmark 'Extract' not found in {Path.GetFileName(filePath)}.");

            // Get the text between the bookmark markers.
            string extractedText = bookmark.Text;

            // Create a new document to hold the extracted content.
            Document resultDoc = new Document();
            DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);
            resultBuilder.Writeln(extractedText);

            // Save the result as PDF.
            string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + "_Extracted.pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            resultDoc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }
    }
}
