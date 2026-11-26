using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font for any content that will be inserted next.
        builder.Font.Name = "Arial";

        // Insert a paragraph of text that will use the default font.
        builder.Writeln("This text uses the default font set via DocumentBuilder.Font.Name.");

        // Define the output file path.
        string outputPath = "DefaultFontExample.docx";

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            Console.WriteLine("Error: The output file was not created.");
            return;
        }

        // Load the saved document to validate the font assignment.
        Document loadedDoc = new Document(outputPath);

        // Retrieve the first Run node (the inserted text).
        Run firstRun = (Run)loadedDoc.GetChild(NodeType.Run, 0, true);

        // Get the font name of the first run.
        string appliedFontName = firstRun.Font.Name;

        // Output the validation result.
        Console.WriteLine($"The font of the first run is: {appliedFontName}");
    }
}
