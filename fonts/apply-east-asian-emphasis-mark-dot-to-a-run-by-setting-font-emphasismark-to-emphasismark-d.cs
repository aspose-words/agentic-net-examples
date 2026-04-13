using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply an emphasis mark (dot) to the text.
        // In Aspose.Words the closest representation is OverSolidCircle.
        builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle;

        // Write sample text with the emphasis mark applied.
        builder.Write("Text with emphasis mark");

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmphasisMarkExample.docx");
        doc.Save(outputPath);

        // Validate that the emphasis mark was set correctly.
        if (builder.Font.EmphasisMark != EmphasisMark.OverSolidCircle)
            throw new InvalidOperationException("Emphasis mark was not applied as expected.");

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
