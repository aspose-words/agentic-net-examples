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

        // Apply an East Asian emphasis mark. The enum does not contain 'Dot',
        // so we use the closest available value: OverSolidCircle.
        builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle;

        // Write sample text that will display the emphasis mark.
        builder.Write("East Asian emphasis mark: OverSolidCircle");

        // Validate that the emphasis mark was set correctly.
        if (builder.Font.EmphasisMark != EmphasisMark.OverSolidCircle)
            throw new InvalidOperationException("EmphasisMark was not set to OverSolidCircle.");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmphasisMarkOverSolidCircle.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
