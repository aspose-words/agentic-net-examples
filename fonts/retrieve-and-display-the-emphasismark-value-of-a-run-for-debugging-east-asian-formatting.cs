using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a paragraph and a run with some East Asian text.
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "東アジアのテキスト");

        // Set an East Asian emphasis mark for debugging purposes.
        // The EmphasisMark enum may not be available in older versions of Aspose.Words,
        // so we set the property via reflection using the underlying integer value.
        // In the current API, the value '0' corresponds to EmphasisMark.None.
        PropertyInfo emphasisProp = typeof(Font).GetProperty("EmphasisMark");
        if (emphasisProp != null && emphasisProp.CanWrite)
        {
            // 0 = None
            emphasisProp.SetValue(run.Font, 0);
        }

        // Add the run to the paragraph and the paragraph to the document.
        paragraph.AppendChild(run);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Retrieve and display the EmphasisMark value.
        object emphasisValue = null;
        if (emphasisProp != null && emphasisProp.CanRead)
        {
            emphasisValue = emphasisProp.GetValue(run.Font);
        }

        Console.WriteLine($"EmphasisMark value (numeric): {emphasisValue}");

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
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
