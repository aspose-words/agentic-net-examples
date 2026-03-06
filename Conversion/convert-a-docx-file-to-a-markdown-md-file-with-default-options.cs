using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToMarkdownConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\sample.docx";

        // Path where the resulting Markdown file will be saved.
        string outputPath = @"C:\Docs\sample.md";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Create default Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Save the document as Markdown using the specified options.
        doc.Save(outputPath, saveOptions);

        Console.WriteLine("Conversion completed successfully.");
    }
}
