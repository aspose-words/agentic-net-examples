using System;
using Aspose.Words;

class ConvertDot
{
    static void Main()
    {
        // Path to the source DOT (template) file.
        string inputPath = "Template.dot";

        // Path to the desired output file (e.g., DOCX format).
        string outputPath = "Converted.docx";

        // Load the DOT document. The format is detected automatically.
        Document doc = new Document(inputPath);

        // Save the document in the chosen format.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
