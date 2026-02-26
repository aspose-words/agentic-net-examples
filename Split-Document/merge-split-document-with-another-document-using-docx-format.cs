using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the first split document.
        Document part1 = new Document("Part1.docx");

        // Load the second document that should be merged.
        Document part2 = new Document("Part2.docx");

        // Append the second document to the end of the first one.
        // KeepSourceFormatting preserves the original styles of part2.
        part1.AppendDocument(part2, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document in DOCX format.
        part1.Save("MergedDocument.docx");
    }
}
