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

        // Load the DOCM document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Save the document in Markdown format. The Save method determines the format from the SaveFormat enum.
        doc.Save(outputPath, SaveFormat.Markdown);
    }
}
