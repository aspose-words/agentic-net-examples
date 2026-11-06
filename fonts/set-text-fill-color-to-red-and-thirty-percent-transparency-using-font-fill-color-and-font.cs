using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with sample text.
        builder.Writeln("Sample text with red fill and 30% transparency.");

        // Retrieve the first Run node that contains the text.
        Run run = (Run)doc.GetChild(NodeType.Run, 0, true);

        // Create Aspose.Drawing.Color red and convert to System.Drawing.Color.
        Aspose.Drawing.Color asposeRed = Aspose.Drawing.Color.Red;
        System.Drawing.Color sysRed = System.Drawing.Color.FromArgb(asposeRed.ToArgb());

        // Apply fill color and transparency to the font.
        run.Font.Fill.Color = sysRed;
        run.Font.Fill.Transparency = 0.3; // 30% transparency

        // Validate that the properties were set correctly.
        bool colorMatches = run.Font.Fill.Color.ToArgb() == sysRed.ToArgb();
        bool transparencyMatches = Math.Abs(run.Font.Fill.Transparency - 0.3) < 0.0001;

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file exists and properties are correct.
        if (System.IO.File.Exists(outputPath) && colorMatches && transparencyMatches)
        {
            Console.WriteLine("Document created successfully with the specified font fill.");
        }
        else
        {
            Console.WriteLine("Document creation failed or properties not set correctly.");
        }
    }
}
