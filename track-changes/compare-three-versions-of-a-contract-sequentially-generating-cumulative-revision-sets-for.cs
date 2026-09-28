using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create version 1 of the contract.
        Document docV1 = new Document();
        DocumentBuilder builderV1 = new DocumentBuilder(docV1);
        builderV1.Writeln("Contract Version 1");
        builderV1.Writeln("This contract is between Party A and Party B.");
        docV1.Save("Contract_v1.docx");

        // Create version 2 with revisions made from version 1.
        Document docV2 = new Document("Contract_v1.docx");
        docV2.StartTrackRevisions("Editor2", DateTime.Now);
        DocumentBuilder builderV2 = new DocumentBuilder(docV2);
        builderV2.Writeln("Additional Clause: Party A shall deliver goods by end of Q4.");
        docV2.StopTrackRevisions();
        docV2.Save("Contract_v2.docx");

        // Create version 3 with revisions made from version 2.
        Document docV3 = new Document("Contract_v2.docx");
        docV3.StartTrackRevisions("Editor3", DateTime.Now);
        // Delete the first paragraph.
        Paragraph firstParagraph = docV3.FirstSection.Body.Paragraphs[0];
        firstParagraph.Remove();
        // Change formatting of the new first paragraph.
        Paragraph secondParagraph = docV3.FirstSection.Body.Paragraphs[0];
        secondParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        docV3.StopTrackRevisions();
        docV3.Save("Contract_v3.docx");

        // Comparison 1: revisions between version 1 and version 2.
        Document revDocV2 = new Document("Contract_v2.docx");
        Console.WriteLine("Revisions between v1 and v2:");
        foreach (Revision rev in revDocV2.Revisions)
        {
            Console.WriteLine($"{rev.RevisionType} by {rev.Author} at {rev.DateTime}");
        }

        // Comparison 2: revisions between version 2 and version 3.
        Document revDocV3 = new Document("Contract_v3.docx");
        Console.WriteLine("Revisions between v2 and v3:");
        foreach (Revision rev in revDocV3.Revisions)
        {
            Console.WriteLine($"{rev.RevisionType} by {rev.Author} at {rev.DateTime}");
        }

        // Cumulative revisions from version 1 through version 3.
        Console.WriteLine("Cumulative revisions from v1 to v3:");
        foreach (Revision rev in revDocV3.Revisions)
        {
            Console.WriteLine($"{rev.RevisionType} by {rev.Author} at {rev.DateTime}");
        }
    }
}
