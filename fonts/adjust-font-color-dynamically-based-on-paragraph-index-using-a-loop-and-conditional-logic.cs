using System;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs to the document.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Loop through each paragraph and set its font color based on the paragraph index.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];

            // Choose a color: even index -> Red, odd index -> Blue.
            Aspose.Drawing.Color aspColor = (i % 2 == 0) ? Aspose.Drawing.Color.Red : Aspose.Drawing.Color.Blue;

            // Convert Aspose.Drawing.Color to System.Drawing.Color as required by Font.Color.
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

            // Apply the color to all runs (text portions) in the paragraph.
            foreach (Run run in para.Runs)
            {
                run.Font.Color = sysColor;
            }
        }

        // Save the document.
        string outputPath = "DynamicFontColors.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
