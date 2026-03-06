using System;
using Aspose.Words;

public class DocumentTextExtractor
{
    // Retrieves the full text of a DOCX document using the document's Range.Text property.
    public static string GetDocumentText(string docxFilePath)
    {
        // Load the DOCX file (lifecycle rule: use Document constructor for loading).
        Document doc = new Document(docxFilePath);

        // Access the Range of the whole document and get its text.
        string text = doc.Range.Text;

        // Return the extracted text.
        return text;
    }

    // Example usage.
    public static void Main()
    {
        string path = "InputDocument.docx"; // Replace with your DOCX file path.
        string content = GetDocumentText(path);
        Console.WriteLine("Document Text:");
        Console.WriteLine(content);
    }
}
