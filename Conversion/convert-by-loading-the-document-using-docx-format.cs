using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string inputPath = @"C:\Docs\Sample.docx";

        // Configure load options to explicitly specify the DOCX format.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Docx // Ensure the document is interpreted as DOCX.
        };

        // Load the document using the specified load options.
        Document doc = new Document(inputPath, loadOptions);

        // Example operation: output the number of pages in the loaded document.
        Console.WriteLine($"Document loaded successfully. Page count: {doc.PageCount}");

        // (Optional) Save the document to another format, e.g., PDF.
        // doc.Save(@"C:\Docs\Sample.pdf");
    }
}
