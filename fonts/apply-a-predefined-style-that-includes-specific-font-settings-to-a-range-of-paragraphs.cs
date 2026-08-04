using System;
using System.IO;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Drawing;
using Aspose.Words.Themes;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom paragraph style named "MyCustomStyle".
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Define specific font settings for the style.
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        // Create a color using Aspose.Drawing, then convert to System.Drawing.Color.
        Aspose.Drawing.Color aspColor = Aspose.Drawing.Color.Blue;
        System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

        customStyle.Font.Color = sysColor;

        // Apply the custom style to a range of paragraphs.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("Paragraph 1 with custom style.");
        builder.Writeln("Paragraph 2 with custom style.");
        builder.Writeln("Paragraph 3 with custom style.");

        // Switch back to the default "Normal" style for subsequent paragraphs.
        builder.ParagraphFormat.Style = doc.Styles["Normal"];
        builder.Writeln("Paragraph 4 with normal style.");

        // Validate that the style's font properties are set correctly.
        Debug.Assert(customStyle.Font.Name == "Arial");
        Debug.Assert(customStyle.Font.Size == 14);
        Debug.Assert(customStyle.Font.Bold == true);
        Debug.Assert(customStyle.Font.Color.ToArgb() == sysColor.ToArgb());

        // Save the document to a file.
        string outputPath = "StyledParagraphs.docx";
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
