using System;
using System.IO;
using Aspose.Words;

public class DocumentSplitter
{
    /// <summary>
    /// Splits the input DOCX document into separate DOCX files, one per section.
    /// </summary>
    /// <param name="inputFilePath">Full path to the source document.</param>
    /// <param name="outputFolderPath">Folder where the split documents will be saved.</param>
    public static void SplitDocumentBySection(string inputFilePath, string outputFolderPath)
    {
        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolderPath))
            Directory.CreateDirectory(outputFolderPath);

        // Load the source document (create/load rule).
        Document sourceDoc = new Document(inputFilePath);

        // Iterate through each section in the source document.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document to hold the current section.
            Document partDoc = new Document();

            // Import the current section into the new document.
            // ImportNode clones the node and its children, preserving formatting.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);
            partDoc.Sections.Add(importedSection);

            // Build the output file name (e.g., Part_1.docx, Part_2.docx, ...).
            string outputFilePath = Path.Combine(outputFolderPath, $"Part_{i + 1}.docx");

            // Save the part as a DOCX file (save rule).
            partDoc.Save(outputFilePath, SaveFormat.Docx);
        }
    }

    // Example usage.
    public static void Main()
    {
        string inputPath = @"C:\Docs\SourceDocument.docx";
        string outputPath = @"C:\Docs\SplitParts";

        SplitDocumentBySection(inputPath, outputPath);

        Console.WriteLine("Document split completed.");
    }
}
