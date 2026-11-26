using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with sample text.
        builder.Writeln("Sample text for font fill color and transparency.");

        // Retrieve the first Run node that contains the text.
        Run run = (Run)doc.GetChildNodes(NodeType.Run, true)[0];

        // Set the fill color to red using Aspose.Drawing.Color,
        // converting it to System.Drawing.Color as required by the API.
        run.Font.Fill.Color = System.Drawing.Color.FromArgb(Color.Red.ToArgb());

        // Set the fill transparency to 30% (0.3).
        run.Font.Fill.Transparency = 0.3f;

        // Validate that the properties were set correctly.
        bool colorMatches = run.Font.Fill.Color.ToArgb() == System.Drawing.Color.FromArgb(Color.Red.ToArgb()).ToArgb();
        bool transparencyMatches = Math.Abs(run.Font.Fill.Transparency - 0.3f) < 0.0001f;

        if (!colorMatches || !transparencyMatches)
            throw new Exception("Font fill properties were not set correctly.");

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!System.IO.File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
