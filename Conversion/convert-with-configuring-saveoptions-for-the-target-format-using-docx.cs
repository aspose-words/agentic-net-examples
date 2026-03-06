using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocxWithOptions
{
    static void Main()
    {
        // Paths to the source and destination files.
        string inputFile = @"C:\Data\SourceDocument.docx";
        string outputFile = @"C:\Data\ConvertedDocument.docx";

        // Load the existing document.
        Document doc = new Document(inputFile);

        // Create a SaveOptions object suitable for the DOCX format.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);

        // Example configuration: enable pretty formatting of the output XML.
        saveOptions.PrettyFormat = true;

        // Save the document using the configured options.
        doc.Save(outputFile, saveOptions);
    }
}
