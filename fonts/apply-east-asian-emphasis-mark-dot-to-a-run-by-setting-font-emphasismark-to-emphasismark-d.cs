using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply an East Asian emphasis mark. The Aspose.Words EmphasisMark enum does not contain a
        // 'Dot' value; the closest equivalent is OverSolidCircle.
        builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle;

        // Write sample text that will display the emphasis mark.
        builder.Write("Text with East Asian emphasis mark (OverSolidCircle)");

        // Save the document.
        const string outputPath = "EmphasisMarkOverSolidCircle.docx";
        doc.Save(outputPath);
    }
}
