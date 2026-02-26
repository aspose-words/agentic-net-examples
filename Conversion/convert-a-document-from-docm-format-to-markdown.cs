using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the macro-enabled DOCM document from the file system.
        Document doc = new Document("input.docm");

        // Save the loaded document in Markdown format.
        // The SaveFormat enumeration value 'Markdown' specifies the target format.
        doc.Save("output.md", SaveFormat.Markdown);
    }
}
