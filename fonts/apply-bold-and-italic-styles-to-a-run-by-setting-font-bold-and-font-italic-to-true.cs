using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply bold and italic formatting to the current font.
        Aspose.Words.Font font = builder.Font;
        font.Bold = true;
        font.Italic = true;

        // Insert a line of text that will inherit the formatting.
        builder.Writeln("This text is bold and italic.");

        // Validate that the formatting was applied.
        if (!font.Bold || !font.Italic)
        {
            throw new InvalidOperationException("Failed to set font styles.");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BoldItalicRun.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved.", outputPath);
        }
    }
}
