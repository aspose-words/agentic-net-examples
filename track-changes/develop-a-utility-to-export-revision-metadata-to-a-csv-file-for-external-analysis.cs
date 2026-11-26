using System;
using System.IO;
using System.Text;
using Aspose.Words;

namespace RevisionExportUtility
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and add initial content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the original paragraph.");

            // Enable track changes.
            string author = "SampleAuthor";
            DateTime revisionDate = DateTime.Now;
            doc.StartTrackRevisions(author, revisionDate);

            // 1. Insert a new paragraph while tracking.
            builder.Writeln("This paragraph was inserted while tracking changes.");

            // 2. Delete the original paragraph while tracking.
            Paragraph originalParagraph = doc.FirstSection.Body.Paragraphs[0];
            originalParagraph.Remove();

            // 3. Change formatting of the inserted paragraph while tracking.
            Paragraph insertedParagraph = doc.FirstSection.Body.Paragraphs[0];
            insertedParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            insertedParagraph.Runs[0].Font.Size = 16;

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Export revision metadata to CSV.
            string csvPath = "revisions.csv";
            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                // Write CSV header.
                writer.WriteLine("RevisionType,Author,DateTime,NodeType,Text");

                // Iterate through all revisions.
                foreach (Revision rev in doc.Revisions)
                {
                    string revType = rev.RevisionType.ToString();
                    string revAuthor = rev.Author;
                    // Use the correct property for revision date/time.
                    string revDate = rev.DateTime.ToString("o"); // ISO 8601 format
                    string nodeType = rev.ParentNode.NodeType.ToString();

                    // Get the text associated with the revision, if any.
                    string text = string.Empty;
                    if (rev.ParentNode is Run run)
                    {
                        text = run.Text;
                    }
                    else if (rev.ParentNode is Paragraph para)
                    {
                        text = para.GetText().TrimEnd('\r', '\n');
                    }

                    // Escape commas in text.
                    if (text.Contains(","))
                        text = $"\"{text}\"";

                    writer.WriteLine($"{revType},{revAuthor},{revDate},{nodeType},{text}");
                }
            }

            // Accept all revisions and save the final document.
            doc.AcceptAllRevisions();
            doc.Save("final_document.docx");
        }
    }
}
