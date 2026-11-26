using System;
using Aspose.Words;

namespace AsposeWordsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // ---------- First section ----------
            // Create a section and its body, then add to the document.
            Section section1 = new Section(doc);
            Body body1 = new Body(doc);
            section1.AppendChild(body1);
            doc.Sections.Add(section1);

            // Add a paragraph with a run to the first section's body.
            Paragraph para1 = new Paragraph(doc);
            para1.AppendChild(new Run(doc, "This is the first section."));
            body1.AppendChild(para1);

            // ---------- Second section ----------
            // Create a second section with its body.
            Section section2 = new Section(doc);
            Body body2 = new Body(doc);
            section2.AppendChild(body2);
            doc.Sections.Add(section2);

            // Add a paragraph with a run to the second section's body.
            Paragraph para2 = new Paragraph(doc);
            para2.AppendChild(new Run(doc, "This is the second section."));
            body2.AppendChild(para2);

            // Save the document to a local file.
            string filePath = "SampleDocument.docx";
            doc.Save(filePath);

            // Load the document from the saved file.
            Document loadedDoc = new Document(filePath);

            // Iterate through each section and extract its plain text using Section.Range.Text.
            for (int i = 0; i < loadedDoc.Sections.Count; i++)
            {
                Section sec = loadedDoc.Sections[i];
                // Trim removes the trailing newline that Aspose.Words adds after each section.
                string sectionText = sec.Range.Text.Trim();
                Console.WriteLine($"Section {i + 1} text: {sectionText}");
            }
        }
    }
}
