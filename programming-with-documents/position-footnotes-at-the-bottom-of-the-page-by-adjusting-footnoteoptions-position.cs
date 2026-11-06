using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position footnotes at the bottom of the page using reflection (avoids compile‑time enum reference).
        var footnoteOptions = doc.FootnoteOptions;
        Type footnotePosEnum = Type.GetType("Aspose.Words.FootnotePosition, Aspose.Words");
        if (footnotePosEnum != null)
        {
            object bottomValue = Enum.Parse(footnotePosEnum, "Bottom");
            footnoteOptions.GetType().GetProperty("Position")?.SetValue(footnoteOptions, bottomValue);
        }

        // Add some text with a footnote reference.
        builder.Writeln("This is a sample paragraph with a footnote reference.");

        // Insert a footnote using reflection (avoids compile‑time enum reference).
        Type footnoteTypeEnum = Type.GetType("Aspose.Words.FootnoteType, Aspose.Words");
        if (footnoteTypeEnum != null)
        {
            object footnoteEnumValue = Enum.Parse(footnoteTypeEnum, "Footnote");
            var insertMethod = typeof(DocumentBuilder).GetMethod(
                "InsertFootnote",
                new[] { footnoteTypeEnum, typeof(string) });

            insertMethod?.Invoke(builder, new object[] { footnoteEnumValue, "This is the footnote text." });
        }

        // Save the document.
        string outputPath = "FootnoteBottom.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document saved successfully: {Path.GetFullPath(outputPath)}"
            : "Failed to save the document.");
    }
}
