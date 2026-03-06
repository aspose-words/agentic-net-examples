using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = "Document.doc";

        // Path for the resulting Markdown file.
        string targetPath = "Document.md";

        // Load the DOC document.
        Document doc = new Document(sourcePath);

        // Save the document in Markdown format.
        doc.Save(targetPath, SaveFormat.Markdown);
    }
}
