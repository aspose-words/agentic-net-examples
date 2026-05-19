using Aspose.Words;
using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (created by default).
        var paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create the first run and add it to the paragraph.
        Run run1 = new Run(doc, "First run. ");
        paragraph.AppendChild(run1);

        // Obtain the Font object from the first run.
        Aspose.Words.Font sharedFont = run1.Font;

        // Set the desired font size.
        sharedFont.Size = 24;

        // Create an Aspose.Drawing.Color (e.g., Blue) and convert it to System.Drawing.Color.
        Aspose.Drawing.Color asposeColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(asposeColor.ToArgb());

        // Assign the color to the font.
        sharedFont.Color = sysColor;

        // Create additional runs.
        Run run2 = new Run(doc, "Second run. ");
        Run run3 = new Run(doc, "Third run.");

        // Append the runs to the paragraph.
        paragraph.AppendChild(run2);
        paragraph.AppendChild(run3);

        // Apply the same formatting to the other runs.
        run2.Font.Size = sharedFont.Size;
        run2.Font.Color = sharedFont.Color;

        run3.Font.Size = sharedFont.Size;
        run3.Font.Color = sharedFont.Color;

        // Validate that all runs share the same size and color.
        bool consistent = run1.Font.Size == run2.Font.Size &&
                          run2.Font.Size == run3.Font.Size &&
                          run1.Font.Color.ToArgb() == run2.Font.Color.ToArgb() &&
                          run2.Font.Color.ToArgb() == run3.Font.Color.ToArgb();

        // Save the document.
        string outputPath = "FontRunsExample.docx";
        doc.Save(outputPath);

        // Ensure the file was created and formatting is consistent.
        if (!File.Exists(outputPath) || !consistent)
            throw new Exception("Document validation failed.");
    }
}
