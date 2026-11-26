using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputEpubs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files with heading styles to represent chapters.
        CreateSamplePdf(Path.Combine(inputFolder, "Sample1.pdf"), "Sample Document 1");
        CreateSamplePdf(Path.Combine(inputFolder, "Sample2.pdf"), "Sample Document 2");

        // Batch convert each PDF to EPUB, preserving the heading structure.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfPath in pdfFiles)
        {
            Document doc = new Document(pdfPath);
            string epubFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".epub";
            string epubPath = Path.Combine(outputFolder, epubFileName);
            doc.Save(epubPath, SaveFormat.Epub);

            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
        }
    }

    private static void CreateSamplePdf(string filePath, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title as Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln(title);

        // Chapter 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter one.");

        // Chapter 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter two.");

        doc.Save(filePath, SaveFormat.Pdf);

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"PDF file was not created: {filePath}");
    }
}
