using System;
using System.IO;

public class Program
{
    // Path to the audit log file.
    private const string AuditFilePath = "audit.log";

    // Entry point of the console application.
    public static void Main()
    {
        // Ensure any existing audit file is removed for a clean run.
        if (File.Exists(AuditFilePath))
        {
            File.Delete(AuditFilePath);
        }

        // Simulate a few cancellation events.
        LogCancellation("User requested cancellation of operation X.");
        System.Threading.Thread.Sleep(500); // Simulate time gap.
        LogCancellation("System timeout caused cancellation of operation Y.");

        // Verify that the audit file was created.
        if (!File.Exists(AuditFilePath))
        {
            throw new InvalidOperationException("Audit file was not created.");
        }

        // Optionally, display the contents (not required, but harmless).
        // Console.WriteLine(File.ReadAllText(AuditFilePath));
    }

    // Writes a cancellation event with a timestamp to the audit file.
    private static void LogCancellation(string reason)
    {
        string timestamp = DateTime.UtcNow.ToString("o"); // ISO 8601 format.
        string logEntry = $"[{timestamp}] Cancellation: {reason}";

        // Append the log entry to the audit file.
        File.AppendAllText(AuditFilePath, logEntry + Environment.NewLine);
    }
}
