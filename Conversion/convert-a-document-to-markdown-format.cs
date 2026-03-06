using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class MarkdownConverter
{
    // Converts a document from any supported format to Markdown.
    public static void ConvertToMarkdown(string inputFilePath, string outputFilePath)
    {
        // Load the source document.
        Document doc = new Document(inputFilePath);

        // Create save options for Markdown output.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Explicitly set the format to Markdown (required by the options object).
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document using the Markdown options.
        doc.Save(outputFilePath, saveOptions);
    }

    // Example entry point.
    public static void Main()
    {
        // Adjust these paths as needed.
        string inputPath = "MyDir/Document.docx";
        string outputPath = "ArtifactsDir/Document.md";

        ConvertToMarkdown(inputPath, outputPath);
    }
}
