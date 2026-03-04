using System;
using Aspose.Words;

class ConvertToHtml
{
    static void Main()
    {
        // Path to the source document. Aspose.Words will auto‑detect the format (DOC, DOCX, RTF, etc.).
        string inputPath = "input.docx";

        // Desired path for the HTML output.
        string outputPath = "output.html";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Save the loaded document in HTML format.
        doc.Save(outputPath, SaveFormat.Html);
    }
}
