using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        var doc = new Aspose.Words.Document();
        var builder = new Aspose.Words.DocumentBuilder(doc);

        // Add a paragraph with a custom line spacing (exactly 30 points).
        builder.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Exactly;
        builder.ParagraphFormat.LineSpacing = 30;
        builder.Writeln("Paragraph with exact line spacing of 30 points.");

        // Add another paragraph with a different custom line spacing (at least 20 points).
        builder.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.AtLeast;
        builder.ParagraphFormat.LineSpacing = 20;
        builder.Writeln("Paragraph with at‑least line spacing of 20 points.");

        // Reset line spacing for all paragraphs to the default value.
        // In Aspose.Words the default line spacing is represented by a value of 0.
        var paragraphs = doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, true);
        foreach (Aspose.Words.Paragraph para in paragraphs)
        {
            para.ParagraphFormat.LineSpacing = 0;
            // Reset the line spacing rule to the default (Multiple) to ensure default behavior.
            para.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Multiple;
        }

        // Save the modified document.
        const string outputPath = "ResetLineSpacing.docx";
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            // File exists – nothing else to do.
        }
    }
}
