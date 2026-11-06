using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the sample DOCX file.
        string docPath = "Sample.docx";

        // Create a sample DOCX file if it does not already exist.
        if (!File.Exists(docPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words!");
            doc.Save(docPath);
        }

        // Load the DOCX file into a Document object.
        Document loadedDoc = new Document(docPath);

        // Validate that the document was loaded and contains at least one page.
        if (loadedDoc.PageCount < 1)
        {
            throw new InvalidOperationException("The loaded document contains no pages.");
        }

        // Output a simple confirmation.
        Console.WriteLine($"Document loaded successfully. Page count: {loadedDoc.PageCount}");
    }
}
