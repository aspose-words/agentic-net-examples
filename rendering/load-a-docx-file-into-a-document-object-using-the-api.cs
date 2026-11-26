using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the sample DOCX file.
        string docxPath = "Sample.docx";

        // Create a sample DOCX document if it does not already exist.
        if (!File.Exists(docxPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words!");
            doc.Save(docxPath);
        }

        // Load the DOCX file into a Document object.
        Document loadedDocument = new Document(docxPath);

        // Simple verification: output the number of pages in the loaded document.
        Console.WriteLine($"Document loaded successfully. Page count: {loadedDocument.PageCount}");
    }
}
