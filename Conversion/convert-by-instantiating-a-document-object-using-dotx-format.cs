using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOTX template file.
        string dotxPath = "Template.dotx";

        // Load the DOTX file into a Document object using the constructor that accepts a file name.
        Document doc = new Document(dotxPath);

        // Example modification: add a paragraph to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document was loaded from a DOTX template and saved as DOCX.");

        // Save the document in DOCX format (or any other desired format).
        string outputPath = "Converted.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
