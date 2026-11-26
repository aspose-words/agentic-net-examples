using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font size for all subsequently inserted text.
        builder.Font.Size = 16; // 16 points

        // Validate that the font size was applied.
        if (builder.Font.Size != 16)
        {
            throw new InvalidOperationException("Failed to set the default font size.");
        }

        // Insert paragraphs that will use the default font size.
        builder.Writeln("First paragraph with the default font size.");
        builder.Writeln("Second paragraph also using the default font size.");

        // Save the document to disk.
        string outputPath = "DefaultFontSize.docx";
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Indicate successful completion.
        Console.WriteLine("Document created successfully at: " + Path.GetFullPath(outputPath));
    }
}
