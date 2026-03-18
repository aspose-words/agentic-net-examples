using System;
using Aspose.Words;

class ExtractSectionTexts
{
    static void Main()
    {
        // Load the Word document from a file.
        // Replace the path with the actual location of your document.
        string inputFile = "input.docx";
        Document doc = new Document(inputFile);

        // Iterate through each section in the document.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            Section section = doc.Sections[i];

            // Get the plain text of the current section via its Range.Text property.
            string sectionText = section.Range.Text;

            // Output the section index and its trimmed text.
            Console.WriteLine($"--- Section {i + 1} ---");
            Console.WriteLine(sectionText.Trim());
            Console.WriteLine();
        }

        // Optionally, save the document after processing (if any modifications were made).
        // doc.Save("output.docx");
    }
}
