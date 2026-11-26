using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a Run with sample text.
        Run run = new Run(doc, "This text will be blue.");

        // Insert the Run into the document.
        builder.InsertNode(run);

        // Set the Run's font color to blue.
        // Aspose.Words expects System.Drawing.Color, so convert from Aspose.Drawing.Color.
        run.Font.Color = System.Drawing.Color.FromArgb(Color.Blue.ToArgb());

        // Validate that the color was set correctly.
        bool isBlue = run.Font.Color.Equals(System.Drawing.Color.Blue);
        Console.WriteLine("Font color set to blue: " + isBlue);

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        Console.WriteLine("Document saved: " + File.Exists(outputPath));
    }
}
