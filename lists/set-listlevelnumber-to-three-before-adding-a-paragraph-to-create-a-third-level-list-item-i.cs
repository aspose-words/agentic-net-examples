using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on the default template and assign it to the builder.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the list level to 3 (fourth level, because levels are zero‑based: 0,1,2,3).
        // This will make the next paragraph a third‑level list item.
        builder.ListFormat.ListLevelNumber = 3;

        // Add the paragraph that will appear as a third‑level list item.
        builder.Writeln("Third‑level list item");

        // End the list formatting for any subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Prepare an output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "ThirdLevelList.docx");
        doc.Save(outputPath);
    }
}
