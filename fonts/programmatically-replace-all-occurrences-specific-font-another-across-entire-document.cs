using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = "input.docx";

        // Path where the modified document will be saved.
        string outputPath = "output.docx";

        // Font name to be replaced and the replacement font name.
        string oldFontName = "Times New Roman";
        string newFontName = "Arial";

        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Iterate through all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // If the run uses the font we want to replace (case‑insensitive check)...
            if (string.Equals(run.Font.Name, oldFontName, StringComparison.OrdinalIgnoreCase))
            {
                // ...replace it with the new font.
                run.Font.Name = newFontName;
            }
        }

        // Save the updated document.
        doc.Save(outputPath);
    }
}
