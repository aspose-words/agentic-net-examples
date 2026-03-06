using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public static class DocumentConverter
{
    /// <summary>
    /// Converts any supported document format to Markdown.
    /// </summary>
    /// <param name="inputStream">Stream containing the source document.</param>
    /// <param name="outputStream">Stream where the Markdown output will be written.</param>
    public static void ConvertToMarkdown(Stream inputStream, Stream outputStream)
    {
        // Load the source document using generic LoadOptions (auto‑detect format).
        var loadOptions = new LoadOptions();
        Document doc = new Document(inputStream, loadOptions);

        // Configure Markdown save options.
        var saveOptions = new MarkdownSaveOptions
        {
            SaveFormat = SaveFormat.Markdown
        };

        // Save the document as Markdown to the output stream.
        doc.Save(outputStream, saveOptions);
    }
}

class Program
{
    static void Main(string[] args)
    {
        // Expect two arguments: input file path and output file path.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: DocumentConverter <inputFile> <outputFile>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        using (FileStream input = File.OpenRead(inputPath))
        using (FileStream output = File.Create(outputPath))
        {
            DocumentConverter.ConvertToMarkdown(input, output);
        }

        Console.WriteLine($"Converted '{inputPath}' to Markdown at '{outputPath}'.");
    }
}
