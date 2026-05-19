using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a plain‑text content control (inline) into the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Set the title and tag for later identification.
        sdt.Title = "CustomerName";
        sdt.Tag = "customer-name";

        // Add some placeholder text inside the content control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Enter name here"));
        paragraph.AppendChild(sdt);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "ContentControlTitleTag.docx");
        doc.Save(outputPath);

        // Optional: serialize basic info to JSON to demonstrate the required package.
        var info = new
        {
            Title = sdt.Title,
            Tag = sdt.Tag,
            File = outputPath
        };
        string json = JsonConvert.SerializeObject(info, Formatting.Indented);
        File.WriteAllText(Path.Combine(outputDir, "info.json"), json);
    }
}
