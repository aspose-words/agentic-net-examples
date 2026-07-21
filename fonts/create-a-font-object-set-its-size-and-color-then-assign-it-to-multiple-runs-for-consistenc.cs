using System;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Obtain a Font object from a temporary run to configure shared settings.
        Run tempRun = new Run(doc, string.Empty);
        Aspose.Words.Font sharedFont = tempRun.Font;
        sharedFont.Size = 24; // Set font size in points.

        // Create a color using Aspose.Drawing, then convert to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Red;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());
        sharedFont.Color = sysColor; // Apply the color to the shared font.

        // Create three runs with different text.
        Run run1 = new Run(doc, "First run. ");
        Run run2 = new Run(doc, "Second run. ");
        Run run3 = new Run(doc, "Third run.");

        // Apply the shared font settings to each run.
        ApplyFontSettings(run1.Font, sharedFont);
        ApplyFontSettings(run2.Font, sharedFont);
        ApplyFontSettings(run3.Font, sharedFont);

        // Append the runs to the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(run1);
        paragraph.AppendChild(run2);
        paragraph.AppendChild(run3);

        // Save the document to a file.
        string outputPath = "FontRuns.docx";
        doc.Save(outputPath);
    }

    // Helper method to copy font properties from a source Font to a target Font.
    private static void ApplyFontSettings(Aspose.Words.Font target, Aspose.Words.Font source)
    {
        target.Size = source.Size;
        target.Color = source.Color;
        // Additional properties can be copied here if needed.
    }
}
