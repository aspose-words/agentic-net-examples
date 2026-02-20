using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: DocumentConverter <inputPath> <outputPath>");
            return;
        }

        using var input = File.OpenRead(args[0]);
        using var output = File.Create(args[1]);

        var converter = new DocumentConverter();
        converter.Convert(input, output);
        Console.WriteLine("Conversion completed.");
    }
}

public class DocumentConverter
{
    /// <summary>
    /// Converts a document from the input stream to the output stream using the most appropriate format.
    /// The output format is derived from the detected input format; if conversion is not possible,
    /// the document is saved as DOCX.
    /// </summary>
    /// <param name="inputStream">Stream containing the source document. Must be readable and seekable.</param>
    /// <param name="outputStream">Stream where the converted document will be written.</param>
    public void Convert(Stream inputStream, Stream outputStream)
    {
        // Ensure the input stream supports seeking so we can reset its position after detection.
        if (!inputStream.CanSeek)
            throw new ArgumentException("Input stream must support seeking.", nameof(inputStream));

        // Detect the format of the input document.
        var formatInfo = FileFormatUtil.DetectFileFormat(inputStream);

        // Reset the stream position to the beginning for actual loading.
        inputStream.Position = 0;

        // Load the document using default load options.
        var loadOptions = new LoadOptions();
        var document = new Document(inputStream, loadOptions);

        // Determine the appropriate save format based on the detected load format.
        // If conversion is not supported, fall back to DOCX.
        var saveFormat = FileFormatUtil.LoadFormatToSaveFormat(formatInfo.LoadFormat);
        if (saveFormat == SaveFormat.Unknown)
            saveFormat = SaveFormat.Docx;

        // Save the document to the output stream using the chosen format.
        document.Save(outputStream, saveFormat);
    }
}
