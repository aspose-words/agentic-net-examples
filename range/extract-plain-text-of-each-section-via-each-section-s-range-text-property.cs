using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document (contains one empty section by default).
        Document doc = new Document();

        // First section (already exists) – add a paragraph with some text.
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "This is the first section."));
        doc.FirstSection.Body.AppendChild(para1);

        // Add a second section with its own body.
        Section secondSection = new Section(doc);
        Body body2 = new Body(doc);
        secondSection.AppendChild(body2);
        doc.AppendChild(secondSection);

        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "Content of the second section goes here."));
        body2.AppendChild(para2);

        // Add a third section with its own body.
        Section thirdSection = new Section(doc);
        Body body3 = new Body(doc);
        thirdSection.AppendChild(body3);
        doc.AppendChild(thirdSection);

        Paragraph para3 = new Paragraph(doc);
        para3.AppendChild(new Run(doc, "Third section text is this line."));
        body3.AppendChild(para3);

        // Extract and display plain text of each section using Section.Range.Text.
        int index = 1;
        foreach (Section section in doc.Sections)
        {
            string sectionText = section.Range.Text;
            Console.WriteLine($"--- Section {index} ---");
            Console.WriteLine(sectionText);
            index++;
        }
    }
}
