using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = "input.doc";

        // Path where the Markdown file will be saved.
        string outputPath = "output.md";

        // Load the DOC document from the file system.
        Document doc = new Document(inputPath);

        // Save the document in Markdown format.
        // The SaveFormat enumeration value for Markdown is SaveFormat.Markdown.
        doc.Save(outputPath, SaveFormat.Markdown);
    }
}
