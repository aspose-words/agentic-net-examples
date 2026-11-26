using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

public class Program
{
    public static void Main()
    {
        // Paths (adjust as needed)
        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        string excelPath = Path.Combine(Directory.GetCurrentDirectory(), "Template.xlsx");

        // Ensure the folder exists
        if (!Directory.Exists(folderPath))
        {
            Directory.CreateDirectory(folderPath);
        }

        // Create a sample Excel file if it does not exist
        if (!File.Exists(excelPath))
        {
            File.WriteAllBytes(excelPath, new byte[0]); // placeholder empty file
        }

        // Create sample Word documents if none exist
        var docFiles = Directory.GetFiles(folderPath, "*.docx");
        if (docFiles.Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                string samplePath = Path.Combine(folderPath, $"Sample{i}.docx");
                CreateEmptyWordDocument(samplePath);
            }
            docFiles = Directory.GetFiles(folderPath, "*.docx");
        }

        // Start Word application (stub implementation)
        var wordApp = new Word.Application();
        wordApp.Visible = false;

        foreach (var docPath in docFiles)
        {
            Word.Document doc = null;
            try
            {
                object readOnly = false;
                object isVisible = false;
                doc = wordApp.Documents.Open(docPath, ref readOnly, ref isVisible);
                Word.Range range = doc.Range(0, 0);
                range.InlineShapes.AddOLEObject(
                    ClassType: "Excel.Sheet",
                    FileName: excelPath,
                    LinkToFile: false,
                    DisplayAsIcon: false,
                    IconFileName: Type.Missing,
                    IconIndex: Type.Missing,
                    IconLabel: Type.Missing,
                    Range: range);

                doc.Save();
                Console.WriteLine($"Processed document: {Path.GetFileName(docPath)}");
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    // No Marshal.ReleaseComObject needed for the stub implementation
                }
            }
        }

        // Quit Word
        wordApp.Quit();
        // No Marshal.ReleaseComObject needed for the stub implementation
    }

    private static void CreateEmptyWordDocument(string path)
    {
        var wordApp = new Word.Application();
        wordApp.Visible = false;
        Word.Document doc = null;
        try
        {
            doc = wordApp.Documents.Add();
            doc.SaveAs2(path);
            Console.WriteLine($"Created empty document: {Path.GetFileName(path)}");
        }
        finally
        {
            if (doc != null)
            {
                doc.Close();
                // No Marshal.ReleaseComObject needed for the stub implementation
            }
            wordApp.Quit();
        }
    }
}

// ---------------------------------------------------------------------------
// Minimal stub implementation of the Word Interop types used in the example.
// This allows the code to compile and run without requiring the actual COM
// libraries. The stub simply writes diagnostic messages to the console.
// ---------------------------------------------------------------------------
namespace Microsoft.Office.Interop.Word
{
    public class Application
    {
        public bool Visible { get; set; }
        public Documents Documents { get; } = new Documents();

        public void Quit() => Console.WriteLine("Word Application Quit.");
    }

    public class Documents
    {
        public Document Open(string fileName, ref object ReadOnly, ref object Visible)
        {
            Console.WriteLine($"Opening document: {Path.GetFileName(fileName)} (ReadOnly={ReadOnly}, Visible={Visible})");
            return new Document(fileName);
        }

        public Document Add()
        {
            Console.WriteLine("Adding new blank document.");
            return new Document(null);
        }
    }

    public class Document
    {
        private readonly string _path;

        public Document(string path) => _path = path;

        public Range Range(int start, int end) => new Range();

        public InlineShapes InlineShapes { get; } = new InlineShapes();

        public void Save()
        {
            if (!string.IsNullOrEmpty(_path))
            {
                Console.WriteLine($"Saving document: {Path.GetFileName(_path)}");
                // In a real implementation the file would be written here.
            }
        }

        public void SaveAs2(string fileName)
        {
            Console.WriteLine($"Saving new document as: {Path.GetFileName(fileName)}");
            // In a real implementation the file would be created here.
        }

        public void Close()
        {
            Console.WriteLine("Closing document.");
        }
    }

    public class Range
    {
        public InlineShapes InlineShapes { get; } = new InlineShapes();
    }

    public class InlineShapes
    {
        public void AddOLEObject(
            string ClassType,
            string FileName,
            bool LinkToFile,
            bool DisplayAsIcon,
            object IconFileName,
            object IconIndex,
            object IconLabel,
            Range Range)
        {
            Console.WriteLine($"[Stub] Adding OLE object: ClassType={ClassType}, FileName={Path.GetFileName(FileName)}");
        }
    }
}
