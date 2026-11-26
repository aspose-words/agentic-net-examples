using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace AsposeWordsTocExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents at the beginning.
            builder.MoveToDocumentStart();
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Add some headings that will appear in the TOC.
            builder.Writeln(); // Blank line after TOC.

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 2.1");

            // Update fields so the TOC is populated.
            doc.UpdateFields();

            // Save the document (optional, just to have an output file).
            const string outputPath = "SampleWithToc.docx";
            doc.Save(outputPath);

            // Iterate over all fields and extract TOC entries.
            foreach (Field field in doc.Range.Fields)
            {
                // Use the correct enum value for a TOC field.
                if (field.Type == FieldType.FieldTOC)
                {
                    // Cast to FieldToc for easier access.
                    FieldToc tocField = (FieldToc)field;

                    // The TOC entries are located between the field's separator and end nodes.
                    Node separatorNode = tocField.Separator;
                    Node endNode = tocField.End;

                    // Walk through the nodes after the separator until the end node.
                    for (Node curNode = separatorNode.NextSibling; curNode != null && curNode != endNode; curNode = curNode.NextSibling)
                    {
                        if (curNode.NodeType == NodeType.Paragraph)
                        {
                            Paragraph para = (Paragraph)curNode;
                            // Output the plain text of each TOC entry.
                            Console.WriteLine(para.ToString(SaveFormat.Text).Trim());
                        }
                    }
                }
            }
        }
    }
}
