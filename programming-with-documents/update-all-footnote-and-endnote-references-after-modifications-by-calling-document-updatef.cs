using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a footnote reference.
        builder.Writeln("This is a paragraph with a footnote reference.");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote text.");

        // Add a paragraph with an endnote reference.
        builder.Writeln("This is a paragraph with an endnote reference.");
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

        // Modify the footnote and endnote texts.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            Paragraph para = (Paragraph)footnote.FirstChild;
            para.Runs.Clear();

            if (footnote.FootnoteType == FootnoteType.Footnote)
                para.Runs.Add(new Run(doc, "Updated footnote text."));
            else if (footnote.FootnoteType == FootnoteType.Endnote)
                para.Runs.Add(new Run(doc, "Updated endnote text."));
        }

        // Update all fields (including footnote/endnote references).
        doc.UpdateFields();

        // Save the updated document.
        string outputPath = "UpdatedFootnotes.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {Path.GetFullPath(outputPath)}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
