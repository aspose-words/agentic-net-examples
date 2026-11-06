using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph with a custom line spacing value.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "This is a sample paragraph with custom line spacing.");
        // Set a non‑default line spacing via ParagraphFormat.
        para.ParagraphFormat.LineSpacing = 20.0;
        para.AppendChild(run);
        doc.FirstSection.Body.AppendChild(para);

        // Add another paragraph that uses the default spacing.
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "Another paragraph with default spacing."));
        doc.FirstSection.Body.AppendChild(para2);

        // Reset line spacing to the default for all paragraphs.
        // Setting LineSpacing to 0 restores the automatic (default) line‑spacing rule.
        foreach (Paragraph p in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            p.ParagraphFormat.LineSpacing = 0;
            // No need to set LineSpacingRule explicitly; a value of 0 makes it Auto.
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save document.");
        }
    }
}
