using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Apply a preset margin setting (e.g., Narrow) to the first section.
        doc.Sections[0].PageSetup.Margins = Margins.Narrow;

        // Save the document as PDF; the format is inferred from the .pdf extension.
        doc.Save("output.pdf");
    }
}
