using System;
using System.IO;
using Aspose.Words;

public class DocxContentExtractor
{
    /// <summary>
    /// Loads a DOCX file and returns its plain‑text content.
    /// </summary>
    /// <param name="docxPath">Full path to the DOCX file.</param>
    /// <returns>All textual content of the document as a single string.</returns>
    public static string ExtractText(string docxPath)
    {
        // Validate input.
        if (string.IsNullOrEmpty(docxPath))
            throw new ArgumentException("File path cannot be null or empty.", nameof(docxPath));

        if (!File.Exists(docxPath))
            throw new FileNotFoundException($"The file '{docxPath}' does not exist.", docxPath);

        // Load the DOCX document using Aspose.Words.Document.
        Document doc = new Document(docxPath);

        // Document.GetText() returns the concatenated text of the whole document,
        // including paragraph marks (\r\n). Trim the result if you do not need trailing whitespace.
        return doc.GetText();
    }

    // Example usage.
    public static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SampleDocument.docx";

        try
        {
            // Extract the text.
            string content = ExtractText(sourcePath);

            // Output the extracted text to the console.
            Console.WriteLine("Extracted Text:");
            Console.WriteLine(content);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
