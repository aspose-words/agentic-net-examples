using System;
using System.IO;
using Aspose.Words;

public class SectionTextExtractor
{
    /// <summary>
    /// Extracts plain text from a specific section of a DOCX file.
    /// </summary>
    /// <param name="docxPath">Full path to the DOCX document.</param>
    /// <param name="sectionIndex">Zero‑based index of the section to extract.</param>
    /// <returns>Trimmed plain text of the specified section.</returns>
    public static string ExtractSectionText(string docxPath, int sectionIndex)
    {
        // Load the document from file (lifecycle rule: load)
        Document doc = new Document(docxPath);

        // Validate the requested section index
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentOutOfRangeException(nameof(sectionIndex), "Section index is out of range.");

        // Retrieve the desired section
        Section targetSection = doc.Sections[sectionIndex];

        // Get the plain text of the section (Node.GetText)
        string rawText = targetSection.GetText();

        // Return trimmed text
        return rawText.Trim();
    }

    // Example usage
    public static void Main()
    {
        // Path to the input DOCX file
        string inputPath = @"C:\Docs\SampleDocument.docx";

        // Index of the section we want to extract (e.g., second section)
        int sectionIdx = 1;

        // Extract the text
        string sectionText = ExtractSectionText(inputPath, sectionIdx);

        // Output the extracted text to console
        Console.WriteLine("Extracted Section Text:");
        Console.WriteLine(sectionText);

        // Optionally, save the extracted text to a .txt file
        string outputPath = @"C:\Docs\ExtractedSection.txt";
        File.WriteAllText(outputPath, sectionText);
    }
}
