using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving; // SaveFormat resides here in recent versions

public static class DocToPdfConverter
{
    /// <summary>
    /// Converts a DOC/DOCX document supplied as a byte array to a PDF and returns the PDF as a byte array.
    /// No intermediate files are written to disk.
    /// </summary>
    /// <param name="docBytes">Byte array containing the source Word document.</param>
    /// <returns>Byte array containing the generated PDF.</returns>
    public static byte[] ConvertDocToPdf(byte[] docBytes)
    {
        // Load the Word document from the input byte array using a memory stream.
        using (MemoryStream inputStream = new MemoryStream(docBytes))
        {
            Document wordDoc = new Document(inputStream);

            // Prepare an output memory stream to receive the PDF data.
            using (MemoryStream outputStream = new MemoryStream())
            {
                // Save the document to the output stream in PDF format.
                wordDoc.Save(outputStream, SaveFormat.Pdf);
                return outputStream.ToArray();
            }
        }
    }

    // ---------------------------------------------------------------------
    // Entry point required for a console‑application project.
    // ---------------------------------------------------------------------
    public static void Main(string[] args)
    {
        // Example usage: load a DOC file from disk into a byte array, convert, and write PDF to disk.
        // In real scenarios the byte[] could come from a database, web request, etc.
        string sourcePath = "sample.doc"; // replace with an existing DOC/DOCX file path
        string targetPath = "sample.pdf";

        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file '{sourcePath}' not found.");
            return;
        }

        byte[] docBytes = File.ReadAllBytes(sourcePath);
        byte[] pdfBytes = ConvertDocToPdf(docBytes);
        File.WriteAllBytes(targetPath, pdfBytes);
        Console.WriteLine($"PDF generated successfully: {targetPath}");
    }
}
