using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add abbreviations with their expansions.
        builder.Writeln("NASA (National Aeronautics and Space Administration)");
        builder.Writeln("UN (United Nations)");
        builder.Writeln("HTML (HyperText Markup Language)");
        builder.Writeln("CPU (Central Processing Unit)");
        builder.Writeln("RAM (Random Access Memory)");

        // Save the document to a DOCX file.
        doc.Save("Abbreviations.docx");
    }
}
