using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder to hold temporary source documents.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Text that will be written to each source document.
        string[] sourceTexts = { "First source document content.", "Second source document content.", "Third source document content." };
        // Corresponding file names.
        string[] sourceFiles = { Path.Combine(dataDir, "Source0.docx"), Path.Combine(dataDir, "Source1.docx"), Path.Combine(dataDir, "Source2.docx") };

        // Create sample source DOCX files.
        for (int i = 0; i < sourceTexts.Length; i++)
        {
            Document srcDoc = new();
            DocumentBuilder builder = new(srcDoc);
            builder.Writeln(sourceTexts[i]);
            srcDoc.Save(sourceFiles[i]); // Save as DOCX (extension determines format).
        }

        // Destination document that will receive the appended content.
        Document dstDoc = new();

        // Append each source document using a different ImportFormatMode.
        for (int i = 0; i < sourceFiles.Length; i++)
        {
            Document srcDoc = new(sourceFiles[i]); // Load the source document.

            // Choose ImportFormatMode based on the index.
            ImportFormatMode mode = i switch
            {
                0 => ImportFormatMode.UseDestinationStyles,
                1 => ImportFormatMode.KeepSourceFormatting,
                _ => ImportFormatMode.KeepDifferentStyles
            };

            // Append the source document to the destination.
            dstDoc.AppendDocument(srcDoc, mode);
        }

        // Validate that the combined document contains text from all source documents.
        string combinedText = dstDoc.GetText();
        foreach (string txt in sourceTexts)
        {
            if (!combinedText.Contains(txt))
                throw new InvalidOperationException($"Combined document is missing expected text: \"{txt}\"");
        }

        // Output PDF path.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "CombinedOutput.pdf");

        // Save the combined document as PDF.
        dstDoc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new FileNotFoundException("The combined PDF file was not created.", outputPdf);
    }
}
