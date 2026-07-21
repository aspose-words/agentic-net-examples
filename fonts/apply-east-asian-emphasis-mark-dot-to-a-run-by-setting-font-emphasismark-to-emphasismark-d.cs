using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply an East Asian emphasis mark (solid circle) to the current font.
        // The EmphasisMark enumeration does not contain a 'Dot' value; the closest equivalent is OverSolidCircle.
        builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle;

        // Verify that the emphasis mark was set correctly.
        if (builder.Font.EmphasisMark != EmphasisMark.OverSolidCircle)
            throw new InvalidOperationException("Failed to set EmphasisMark to OverSolidCircle.");

        // Write some sample text that will display the emphasis mark.
        builder.Write("East Asian emphasis mark: Dot (represented by OverSolidCircle)");

        // Define the output path and ensure the directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "EmphasisMarkDot.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
