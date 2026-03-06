using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document from the file system.
        Document doc = new Document("input.docx");

        // Save the document in Markdown format.
        // The Save method overload takes the output file name and the desired SaveFormat.
        doc.Save("output.md", SaveFormat.Markdown);
    }
}
