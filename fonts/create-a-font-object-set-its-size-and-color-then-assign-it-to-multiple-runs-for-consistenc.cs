using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Obtain the Font object from the DocumentBuilder.
        var font = builder.Font; // Aspose.Words.Font

        // Set font size.
        font.Size = 14;

        // Set font color using Aspose.Drawing.Color and convert to System.Drawing.Color.
        var aspColor = Aspose.Drawing.Color.Red;
        var sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());
        font.Color = sysColor;

        // Insert multiple runs that share the same font settings.
        builder.Write("First run with custom font. ");
        builder.Write("Second run with the same custom font.");

        // Validate that the font properties were applied to the runs.
        var firstParagraph = doc.FirstSection.Body.FirstParagraph;
        var firstRun = (Run)firstParagraph.Runs[0];
        var secondRun = (Run)firstParagraph.Runs[1];

        if (firstRun.Font.Size != 14 || firstRun.Font.Color.ToArgb() != sysColor.ToArgb())
            throw new System.Exception("Font properties not applied correctly to the first run.");

        if (secondRun.Font.Size != 14 || secondRun.Font.Color.ToArgb() != sysColor.ToArgb())
            throw new System.Exception("Font properties not applied correctly to the second run.");

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!System.IO.File.Exists(outputPath))
            throw new System.Exception("Failed to create the output file.");
    }
}
