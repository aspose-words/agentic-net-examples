using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a working directory relative to the executable location.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsExample");
        Directory.CreateDirectory(workDir);

        // Paths for the source documents.
        string[] sourcePaths = new string[]
        {
            Path.Combine(workDir, "Source1.docx"),
            Path.Combine(workDir, "Source2.docx"),
            Path.Combine(workDir, "Source3.docx")
        };

        // Create three sample source documents with distinct content.
        CreateSampleDocument(sourcePaths[0], "First source document. Style: Normal.", StyleIdentifier.Normal);
        CreateSampleDocument(sourcePaths[1], "Second source document. Style: Heading1.", StyleIdentifier.Heading1);
        CreateSampleDocument(sourcePaths[2], "Third source document. Style: Heading2.", StyleIdentifier.Heading2);

        // Destination document that will receive the appended content.
        Document destination = new Document();

        // Append each source document using a different ImportFormatMode.
        for (int i = 0; i < sourcePaths.Length; i++)
        {
            Document src = new Document(sourcePaths[i]);

            ImportFormatMode mode = i switch
            {
                0 => ImportFormatMode.UseDestinationStyles,
                1 => ImportFormatMode.KeepSourceFormatting,
                _ => ImportFormatMode.KeepDifferentStyles
            };

            destination.AppendDocument(src, mode);
        }

        // Save the combined document as PDF.
        string pdfPath = Path.Combine(workDir, "Combined.pdf");
        destination.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }

        // Optional: clean up source files (comment out if inspection is needed).
        // foreach (var path in sourcePaths) File.Delete(path);
    }

    // Helper method to create a simple DOCX file with specified text and style.
    private static void CreateSampleDocument(string filePath, string text, StyleIdentifier styleId)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = styleId;
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
