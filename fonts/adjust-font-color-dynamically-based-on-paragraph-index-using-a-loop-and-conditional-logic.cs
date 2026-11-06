using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // For Aspose.Drawing.Color
using System.Drawing; // Only for System.Drawing.Color conversion (fully qualified usage)

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs with sample text.
        for (int i = 0; i < 6; i++)
        {
            builder.Writeln($"This is paragraph number {i + 1}.");
        }

        // Loop through all paragraphs and set font color based on the paragraph index.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];

            // Determine color: even index -> Blue, odd index -> Red.
            Aspose.Drawing.Color aspColor = (i % 2 == 0) ? Aspose.Drawing.Color.Blue : Aspose.Drawing.Color.Red;

            // Convert Aspose.Drawing.Color to System.Drawing.Color.
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

            // Apply the color to all runs within the paragraph.
            foreach (Run run in para.Runs)
            {
                run.Font.Color = sysColor;
            }
        }

        // Save the document.
        string outputPath = "DynamicFontColors.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
