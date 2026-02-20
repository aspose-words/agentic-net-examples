using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words;

/// <summary>
/// Demonstrates loading a document of any supported format and saving it as a legacy DOC file.
/// </summary>
public class DocConverter
{
    /// <summary>
    /// Converts the input document to the Microsoft Word 97‑2003 DOC format.
    /// </summary>
    /// <param name="inputPath">Full path to the source document.</param>
    /// <param name="outputPath">Full path where the DOC file will be saved.</param>
    public static void ConvertToDoc(string inputPath, string outputPath)
    {
        // Verify that the source file exists.
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Input file not found: {inputPath}");

        // Detect the format of the input file.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath);

        // Prepare load options with the detected format.
        // This ensures that Aspose.Words uses the correct loader even for ambiguous extensions.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = formatInfo.LoadFormat
        };

        // Load the document using the detected format.
        Document doc = new Document(inputPath, loadOptions);

        // Configure DOC save options (optional – you can tweak properties here).
        DocSaveOptions saveOptions = new DocSaveOptions
        {
            // Example: embed the generator name (default is true).
            ExportGeneratorName = true,
            // Example: compress all metafiles regardless of size.
            AlwaysCompressMetafiles = true
        };

        // Save the document as DOC.
        doc.Save(outputPath, saveOptions);
    }

    // Example usage.
    public static void Main()
    {
        string sourceFile = @"C:\Docs\sample.pdf";   // any supported format
        string targetFile = @"C:\Docs\sample_converted.doc";

        try
        {
            ConvertToDoc(sourceFile, targetFile);
            Console.WriteLine("Conversion succeeded.");
        }
        catch (UnsupportedFileFormatException ex)
        {
            Console.WriteLine($"Unsupported format: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
