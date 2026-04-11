using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ConvertDocxToPdfWithCover
{
    public static void Main()
    {
        // Prepare a working folder.
        string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workFolder);

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image that will be used as the cover page.
        // -----------------------------------------------------------------
        // This is a 1x1 pixel red PNG image (base64 decoded).
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };
        string imagePath = Path.Combine(workFolder, "cover.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // ---------------------------------------------------------------
        // 2. Create a sample DOCX document and insert the cover image.
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move to the very start of the document to place the cover page.
        builder.MoveToDocumentStart();

        // Insert the cover image.
        builder.InsertImage(imagePath);

        // Add a page break after the cover page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add some regular content.
        builder.Writeln("This is the main content of the document.");
        builder.Writeln("It appears after the cover page.");

        // Save the intermediate DOCX file.
        string docxPath = Path.Combine(workFolder, "sample.docx");
        doc.Save(docxPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 3. Load the DOCX file (simulating an existing source) and convert to PDF.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        string pdfPath = Path.Combine(workFolder, "output.pdf");
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------------------------------------------------------------
        // 4. Validate that the PDF file was created successfully.
        // ---------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("The PDF file was not created.", pdfPath);
        }

        FileInfo pdfInfo = new FileInfo(pdfPath);
        if (pdfInfo.Length == 0)
        {
            throw new InvalidOperationException("The PDF file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine($"PDF conversion succeeded. File saved at: {pdfPath}");
    }
}
