using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Add a heading to the destination document.
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Combined Document with Videos");

        // Define source DOCX file paths in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string[] sourceFiles = new string[]
        {
            Path.Combine(baseDir, "VideoDoc1.docx"),
            Path.Combine(baseDir, "VideoDoc2.docx")
        };

        // Ensure each source file exists; if not, create a simple placeholder document.
        foreach (string srcPath in sourceFiles)
        {
            if (!File.Exists(srcPath))
            {
                Document placeholder = new Document();
                DocumentBuilder builder = new DocumentBuilder(placeholder);
                builder.Writeln($"Placeholder content for \"{Path.GetFileName(srcPath)}\".");
                placeholder.Save(srcPath);
            }
        }

        // Append each source document to the destination document.
        foreach (string srcPath in sourceFiles)
        {
            Document srcDoc = new Document(srcPath);
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Prepare PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations,
            UpdateFields = true
        };

        // Ensure the output directory exists.
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the combined document as PDF using the configured options.
        string outputPath = Path.Combine(outputDir, "CombinedWithVideos.pdf");
        dstDoc.Save(outputPath, pdfOptions);

        Console.WriteLine($"Combined PDF saved to: {outputPath}");
    }
}
