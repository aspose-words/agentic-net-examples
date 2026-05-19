using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListAssignmentExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a regular paragraph (not part of any list).
            builder.Writeln("This paragraph is not a list item.");

            // Create a bulleted list that will be reused.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Create a new paragraph that will become a list item.
            Paragraph listParagraph = new Paragraph(doc);
            listParagraph.AppendChild(new Run(doc, "This paragraph belongs to the existing list."));

            // Assign the existing list to the paragraph.
            listParagraph.ListFormat.List = bulletList;
            // Set the list level (0 = first level).
            listParagraph.ListFormat.ListLevelNumber = 0;

            // Append the paragraph to the document body.
            doc.FirstSection.Body.AppendChild(listParagraph);

            // Save the document to disk.
            doc.Save("ListAssignment.docx");
        }
    }
}
