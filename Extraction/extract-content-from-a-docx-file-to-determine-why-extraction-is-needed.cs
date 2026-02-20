using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class DocxTextExtractor
{
    // Extracts plain text from a DOCX file.
    public static string ExtractText(string docxPath)
    {
        // Ensure the file exists.
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("The specified DOCX file was not found.", docxPath);

        try
        {
            // Load the document as plain text. The constructor automatically detects the format.
            PlainTextDocument plainTextDoc = new PlainTextDocument(docxPath);

            // Retrieve the concatenated text content.
            return plainTextDoc.Text;
        }
        catch (FileCorruptedException ex)
        {
            // The document appears to be corrupted.
            Console.Error.WriteLine($"File is corrupted: {ex.Message}");
            throw;
        }
        catch (UnsupportedFileFormatException ex)
        {
            // The file format is not supported by Aspose.Words.
            Console.Error.WriteLine($"Unsupported format: {ex.Message}");
            throw;
        }
    }

    // Example usage.
    static void Main()
    {
        string inputPath = @"C:\Docs\Sample.docx";

        try
        {
            string text = ExtractText(inputPath);
            Console.WriteLine("Extracted Text:");
            Console.WriteLine(text);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Extraction failed: {ex.Message}");
        }
    }
}
