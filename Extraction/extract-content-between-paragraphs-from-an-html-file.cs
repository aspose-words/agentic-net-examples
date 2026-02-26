using System;
using System.IO;
using Aspose.Words;

class ExtractParagraphContent
{
    static void Main()
    {
        // Path to the source HTML file.
        string inputPath = @"C:\Temp\input.html";

        // Path to the output text file that will contain the extracted paragraph contents.
        string outputPath = @"C:\Temp\extracted_paragraphs.txt";

        // Load the HTML document. The Document constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Get the collection of paragraphs from the main body of the first section.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Open a StreamWriter to write the extracted texts.
        using (StreamWriter writer = new StreamWriter(outputPath))
        {
            // Iterate through each paragraph.
            foreach (Paragraph para in paragraphs)
            {
                // Get the text of the paragraph, including the trailing paragraph break character.
                string textWithBreak = para.GetText();

                // Trim the trailing paragraph break characters (\r, \n, or \f) to obtain clean content.
                string cleanText = textWithBreak.TrimEnd('\r', '\n', '\f');

                // Write the cleaned paragraph text to the output file.
                writer.WriteLine(cleanText);
            }
        }

        // Optionally, save the document back to another format (e.g., plain text) if needed.
        // doc.Save(@"C:\Temp\output.txt", SaveFormat.Text);
    }
}
