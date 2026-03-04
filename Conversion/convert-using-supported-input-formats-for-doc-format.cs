using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocumentConverter
{
    // Converts a document from any supported input format to DOC format using file paths.
    public static void ConvertToDoc(string inputFilePath, string outputFilePath)
    {
        // Load the source document. Aspose.Words automatically detects the format.
        Document doc = new Document(inputFilePath);

        // Prepare save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Save the document as DOC.
        doc.Save(outputFilePath, saveOptions);
    }

    // Converts a document from any supported input format to DOC format using streams.
    public static void ConvertToDoc(Stream inputStream, Stream outputStream)
    {
        // Load the source document from the input stream.
        Document doc = new Document(inputStream);

        // Prepare save options for the legacy DOC format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Save the document to the output stream as DOC.
        doc.Save(outputStream, saveOptions);
    }
}

public class Program
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        // Example using file paths.
        string inputPath = "input.pdf"; // TODO: replace with an actual supported input file.
        string outputPath = "output.doc";

        DocumentConverter.ConvertToDoc(inputPath, outputPath);
        Console.WriteLine($"Converted '{inputPath}' to '{outputPath}' using file paths.");

        // Example using streams.
        using (FileStream inStream = File.OpenRead(inputPath))
        using (FileStream outStream = File.Create("output_stream.doc"))
        {
            DocumentConverter.ConvertToDoc(inStream, outStream);
        }
        Console.WriteLine("Converted using streams.");
    }
}
