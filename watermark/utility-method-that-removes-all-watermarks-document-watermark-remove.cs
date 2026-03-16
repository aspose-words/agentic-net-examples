using System;
using Aspose.Words;

public static class WatermarkHelper
{
    /// <summary>
    /// Removes any watermark (text or image) from the specified document and saves the result.
    /// </summary>
    /// <param name="inputFilePath">Path to the source document.</param>
    /// <param name="outputFilePath">Path where the document without watermarks will be saved.</param>
    public static void RemoveAllWatermarks(string inputFilePath, string outputFilePath)
    {
        // Load the document from the file system.
        Document doc = new Document(inputFilePath);

        // If a watermark is present (type is not None), remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the modified document to the desired location.
        doc.Save(outputFilePath);
    }
}

public class Program
{
    /// <summary>
    /// Entry point required for a console application. Demonstrates the helper method.
    /// </summary>
    /// <param name="args">[0] = input file path, [1] = output file path.</param>
    public static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: WatermarkRemover <inputFilePath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            WatermarkHelper.RemoveAllWatermarks(inputPath, outputPath);
            Console.WriteLine($"Watermarks removed successfully. Saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
