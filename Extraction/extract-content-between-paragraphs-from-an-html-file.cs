using System;
using System.IO;
using System.Text;
using Aspose.Words;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Path to the source HTML file.
        string htmlPath = @"C:\Temp\source.html";

        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document(htmlPath);

        // Access the collection of paragraphs in the main body of the document.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Define the range of paragraphs to extract (zero‑based indices).
        // For example, extract content from paragraph 2 up to paragraph 4 inclusive.
        int startIndex = 1; // second paragraph
        int endIndex = 3;   // fourth paragraph

        // Guard against out‑of‑range indices.
        if (startIndex < 0) startIndex = 0;
        if (endIndex >= paragraphs.Count) endIndex = paragraphs.Count - 1;
        if (startIndex > endIndex) throw new ArgumentException("Start index must be less than or equal to end index.");

        // Accumulate the text of the selected paragraphs.
        StringBuilder extractedText = new StringBuilder();
        for (int i = startIndex; i <= endIndex; i++)
        {
            // GetText() returns the paragraph text plus the terminating paragraph break.
            // Trim() removes the trailing break characters.
            extractedText.AppendLine(paragraphs[i].GetText().Trim());
        }

        // Output the extracted content to the console.
        Console.WriteLine("Extracted content between paragraphs:");
        Console.WriteLine(extractedText.ToString());

        // Optionally, save the extracted text to a plain‑text file.
        string outputPath = @"C:\Temp\extracted.txt";
        File.WriteAllText(outputPath, extractedText.ToString());

        // If you also need to save the extracted portion as a new Word document,
        // create a new Document and import the selected paragraphs.
        Document newDoc = new Document();
        newDoc.RemoveAllChildren(); // Ensure the new document is empty.

        // Add a new section and body to host the imported paragraphs.
        Section section = new Section(newDoc);
        newDoc.AppendChild(section);
        Body body = new Body(newDoc);
        section.AppendChild(body);

        // Import each selected paragraph into the new document.
        NodeImporter importer = new NodeImporter(doc, newDoc, ImportFormatMode.KeepSourceFormatting);
        for (int i = startIndex; i <= endIndex; i++)
        {
            Node imported = importer.ImportNode(paragraphs[i], true);
            body.AppendChild(imported);
        }

        // Save the new document containing only the extracted paragraphs.
        string docOutputPath = @"C:\Temp\extracted.docx";
        newDoc.Save(docOutputPath);
    }
}
