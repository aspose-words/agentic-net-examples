using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for temporary source documents and final output.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsWork");
        Directory.CreateDirectory(workDir);

        // Define source documents: file name, text content, and ImportFormatMode to use when appending.
        var sources = new List<(string FileName, string Content, ImportFormatMode Mode)>
        {
            (Path.Combine(workDir, "Doc1.docx"), "First document content.", ImportFormatMode.UseDestinationStyles),
            (Path.Combine(workDir, "Doc2.docx"), "Second document content.", ImportFormatMode.KeepSourceFormatting),
            (Path.Combine(workDir, "Doc3.docx"), "Third document content.", ImportFormatMode.KeepDifferentStyles)
        };

        // Create each source DOCX file.
        foreach (var (fileName, content, _) in sources)
        {
            var srcDoc = new Document();
            var builder = new DocumentBuilder(srcDoc);
            builder.Writeln(content);
            srcDoc.Save(fileName, SaveFormat.Docx);
        }

        // Destination document that will receive all source documents.
        var dstDoc = new Document();

        // Append each source document using its specific ImportFormatMode.
        foreach (var (fileName, _, mode) in sources)
        {
            var srcDoc = new Document(fileName);
            dstDoc.AppendDocument(srcDoc, mode);
        }

        // Save the combined document as PDF.
        string pdfPath = Path.Combine(workDir, "Combined.pdf");
        dstDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file exists.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The combined PDF was not created.");

        // Load the PDF back as a Document to verify its text contains all source contents.
        var pdfDoc = new Document(pdfPath);
        string combinedText = pdfDoc.GetText();

        foreach (var (_, content, _) in sources)
        {
            if (!combinedText.Contains(content))
                throw new InvalidOperationException($"Combined PDF is missing expected content: \"{content}\"");
        }

        // Clean up temporary files (optional).
        // Directory.Delete(workDir, true);
    }
}
