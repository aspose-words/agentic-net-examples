using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font for all subsequently inserted text.
        builder.Font.Name = "Arial";

        // Verify that the font name was set correctly.
        if (builder.Font.Name != "Arial")
            throw new InvalidOperationException("Failed to set default font to Arial.");

        // Insert text using the default font.
        builder.Writeln("This line is formatted with the default font Arial.");

        // Change the default font before adding more text.
        builder.Font.Name = "Times New Roman";

        // Verify the new font setting.
        if (builder.Font.Name != "Times New Roman")
            throw new InvalidOperationException("Failed to change default font to Times New Roman.");

        // Insert additional text with the new default font.
        builder.Writeln("This line uses the default font Times New Roman.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DefaultFontExample.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Confirm that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
