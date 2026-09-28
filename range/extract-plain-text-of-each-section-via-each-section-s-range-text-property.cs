using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document (contains one default section).
        Document doc = new Document();

        // First (default) section already has a Body, add a paragraph to it.
        Section firstSection = doc.FirstSection;
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "This is the first section."));
        firstSection.Body.AppendChild(para1);

        // Create a second section.
        Section secondSection = new Section(doc);
        // A Section must contain a Body node; add it explicitly.
        Body secondBody = new Body(doc);
        secondSection.AppendChild(secondBody);
        // Add a paragraph to the second section's body.
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "This is the second section."));
        secondBody.AppendChild(para2);
        // Add the second section to the document.
        doc.Sections.Add(secondSection);

        // Save the document locally.
        const string fileName = "Sample.docx";
        doc.Save(fileName);

        // Load the document back from the file.
        Document loadedDoc = new Document(fileName);

        // Iterate through each section and output its plain text.
        int sectionIndex = 1;
        foreach (Section section in loadedDoc.Sections)
        {
            string sectionText = section.Range.Text;
            Console.WriteLine($"Section {sectionIndex} text:");
            Console.WriteLine(sectionText);
            sectionIndex++;
        }
    }
}
