using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Save the document in Markdown format.
        // The SaveFormat.Markdown enum value directs Aspose.Words to use the MarkdownSaveOptions internally.
        doc.Save("output.md", SaveFormat.Markdown);
    }
}
