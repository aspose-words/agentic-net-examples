using System;
using Aspose.Words;

class DocmToEpubConverter
{
    static void Main()
    {
        // Path to the folder that contains the input DOCM file.
        string inputFolder = @"C:\Docs\Input";
        // Path to the folder where the EPUB file will be saved.
        string outputFolder = @"C:\Docs\Output";

        // Load the macro‑enabled Word document (DOCM).
        Document doc = new Document(System.IO.Path.Combine(inputFolder, "sample.docm"));

        // Save the document as EPUB. The format is inferred from the ".epub" extension.
        doc.Save(System.IO.Path.Combine(outputFolder, "sample.epub"));
    }
}
