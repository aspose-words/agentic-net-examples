using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class MarkdownExportExample
{
    /// <summary>
    /// Exports a document to Markdown using specific save options.
    /// </summary>
    /// <param name="inputFilePath">Path to the source document.</param>
    /// <param name="outputFilePath">Path where the Markdown file will be saved.</param>
    public static void Export(string inputFilePath, string outputFilePath)
    {
        // Load the source document.
        Document doc = new Document(inputFilePath);

        // Configure Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Export tables as raw HTML.
        saveOptions.ExportAsHtml = MarkdownExportAsHtml.Tables;

        // Export list items as plain text.
        saveOptions.ListExportMode = MarkdownListExportMode.PlainText;

        // Export links using reference style.
        saveOptions.LinkExportMode = MarkdownLinkExportMode.Reference;

        // Include underline formatting using "++".
        saveOptions.ExportUnderlineFormatting = true;

        // Explicitly set the save format to Markdown (optional, kept for clarity).
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document as Markdown with the specified options.
        doc.Save(outputFilePath, saveOptions);
    }
}

public class Program
{
    /// <summary>
    /// Entry point required for a console application.
    /// Usage: dotnet run <inputFilePath> <outputFilePath>
    /// </summary>
    public static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: <program> <inputFilePath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            MarkdownExportExample.Export(inputPath, outputPath);
            Console.WriteLine($"Document exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during export: {ex.Message}");
        }
    }
}
