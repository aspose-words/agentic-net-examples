using System;
using System.IO;
using Aspose.Words;

public class ApplyEmphasisMark
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the East Asian emphasis mark "Dot" (represented by OverSolidCircle) to the current font.
        builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;

        // Validate that the emphasis mark was set correctly.
        if (builder.Font.EmphasisMark != Aspose.Words.EmphasisMark.OverSolidCircle)
            throw new InvalidOperationException("Failed to set EmphasisMark to OverSolidCircle.");

        // Write some sample text that will display the emphasis mark.
        builder.Write("East Asian emphasis mark: Dot");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmphasisMarkDot.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
