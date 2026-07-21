using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a text watermark.
        const string watermarkText = "Sample Watermark";
        doc.Watermark.SetText(watermarkText);

        // Verify that the watermark was added.
        if (doc.Watermark.Type != WatermarkType.Text)
        {
            Console.WriteLine("Error: Text watermark was not added.");
            return;
        }

        // Remove the watermark.
        doc.Watermark.Remove();

        // Verify that the watermark was removed.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            Console.WriteLine("Error: Watermark was not removed.");
        }
        else
        {
            Console.WriteLine("Success: Watermark was removed.");
        }

        // Save the resulting document (optional, demonstrates that the file is still valid).
        const string outputPath = "WatermarkRemovalResult.docx";
        doc.Save(outputPath);
    }
}
