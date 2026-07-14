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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line of text.
        builder.Writeln("Hello, semi‑transparent world!");

        // Retrieve the first run of the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.Paragraphs[0];
        Run run = paragraph.Runs[0];

        // Define a semi‑transparent red fill.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.FromArgb(255, 0, 0); // Red
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Apply fill color and transparency.
        run.Font.Fill.Color = sysColor;
        run.Font.Fill.Transparency = 0.5; // 50 % transparent

        // Validate the applied properties.
        bool colorMatches = run.Font.Fill.Color.ToArgb() == sysColor.ToArgb();
        bool transparencyMatches = Math.Abs(run.Font.Fill.Transparency - 0.5) < 0.0001;

        // Save the document.
        string outputPath = "SemiTransparentText.docx";
        doc.Save(outputPath);

        // Verify that the file was created and properties are set.
        if (File.Exists(outputPath) && colorMatches && transparencyMatches)
        {
            Console.WriteLine("Document created successfully with semi‑transparent text.");
        }
        else
        {
            Console.WriteLine("Document creation failed or properties not set correctly.");
        }
    }
}
