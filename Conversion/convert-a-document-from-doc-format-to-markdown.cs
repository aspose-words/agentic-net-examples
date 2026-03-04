using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = "Document.doc";

        // Path where the Markdown file will be saved.
        string outputFile = "Document.md";

        // Load the DOC document from the file system.
        Document doc = new Document(inputFile);

        // Save the loaded document in Markdown format.
        // The SaveFormat enum value 'Markdown' specifies the target format.
        doc.Save(outputFile, SaveFormat.Markdown);
    }
}
