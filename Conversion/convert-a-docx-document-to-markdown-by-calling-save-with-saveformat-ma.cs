using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the Markdown file will be saved.
        string outputPath = "output.md";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Convert and save the document as Markdown.
        doc.Save(outputPath, SaveFormat.Markdown);
    }
}
