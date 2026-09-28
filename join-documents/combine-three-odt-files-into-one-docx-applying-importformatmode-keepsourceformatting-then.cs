using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folder for sample files
        string folder = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsSample");
        Directory.CreateDirectory(folder);

        // Define source and output file paths
        string doc1Path = Path.Combine(folder, "Doc1.odt");
        string doc2Path = Path.Combine(folder, "Doc2.odt");
        string doc3Path = Path.Combine(folder, "Doc3.odt");
        string outputPath = Path.Combine(folder, "Combined.docx");

        // Create three ODT source documents with distinct content
        CreateSampleDocument(doc1Path, "First document content.");
        CreateSampleDocument(doc2Path, "Second document content.");
        CreateSampleDocument(doc3Path, "Third document content.");

        // Load the first document as the destination
        Document destination = new Document(doc1Path);

        // Append the remaining documents preserving their original formatting
        destination.AppendDocument(new Document(doc2Path), ImportFormatMode.KeepSourceFormatting);
        destination.AppendDocument(new Document(doc3Path), ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as DOCX
        destination.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file was created
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The combined DOCX file was not saved.");

        // Validate that the combined document contains content from all source documents
        Document combined = new Document(outputPath);
        string combinedText = combined.GetText();

        if (!combinedText.Contains("First document content.") ||
            !combinedText.Contains("Second document content.") ||
            !combinedText.Contains("Third document content."))
        {
            throw new InvalidOperationException("The combined document does not contain expected content from all sources.");
        }
    }

    // Helper method to create a simple ODT document with specified text
    private static void CreateSampleDocument(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath, SaveFormat.Odt);
    }
}
