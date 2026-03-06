using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = "input.docm";

        // Path where the Markdown file will be saved.
        string outputPath = "output.md";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Save the document in Markdown format.
        doc.Save(outputPath, SaveFormat.Markdown);
    }
}
