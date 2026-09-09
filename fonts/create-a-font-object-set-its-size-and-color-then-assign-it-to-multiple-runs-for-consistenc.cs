using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Use DocumentBuilder to obtain a Font object.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an Aspose.Drawing.Color and convert it to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        // Configure the shared font.
        Aspose.Words.Font sharedFont = builder.Font;
        sharedFont.Size = 24;
        sharedFont.Color = sysColor;

        // Create the first run and apply the shared font properties.
        Run run1 = new Run(doc, "First run. ");
        run1.Font.Size = sharedFont.Size;
        run1.Font.Color = sharedFont.Color;
        paragraph.AppendChild(run1);

        // Create the second run and apply the same font.
        Run run2 = new Run(doc, "Second run.");
        run2.Font.Size = sharedFont.Size;
        run2.Font.Color = sharedFont.Color;
        paragraph.AppendChild(run2);

        // Validation: ensure both runs have the expected size and color.
        if (run1.Font.Size != 24 || run2.Font.Size != 24)
            throw new InvalidOperationException("Font size mismatch.");

        if (run1.Font.Color.ToArgb() != sysColor.ToArgb() || run2.Font.Color.ToArgb() != sysColor.ToArgb())
            throw new InvalidOperationException("Font color mismatch.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FontRuns.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Output file not found.", outputPath);
    }
}
