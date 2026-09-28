using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Define a shared font.
        var sharedFont = builder.Font;
        sharedFont.Size = 16;

        // Define color using Aspose.Drawing.Color and convert to System.Drawing.Color.
        var awColor = Color.Blue;
        sharedFont.Color = System.Drawing.Color.FromArgb(awColor.ToArgb());

        // Create first run with shared font.
        var run1 = new Run(doc, "Hello, ");
        run1.Font.Size = sharedFont.Size;
        run1.Font.Color = sharedFont.Color;

        // Create second run with shared font.
        var run2 = new Run(doc, "world!");
        run2.Font.Size = sharedFont.Size;
        run2.Font.Color = sharedFont.Color;

        // Add runs to a paragraph.
        var paragraph = new Paragraph(doc);
        paragraph.AppendChild(run1);
        paragraph.AppendChild(run2);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Validate font properties.
        if (run1.Font.Size != 16 || run2.Font.Size != 16)
            throw new Exception("Font size mismatch.");
        var expectedColor = System.Drawing.Color.FromArgb(awColor.ToArgb());
        if (run1.Font.Color.ToArgb() != expectedColor.ToArgb() ||
            run2.Font.Color.ToArgb() != expectedColor.ToArgb())
            throw new Exception("Font color mismatch.");

        // Save document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure file exists.
        if (!File.Exists(outputPath))
            throw new Exception("Output file was not created.");
    }
}
