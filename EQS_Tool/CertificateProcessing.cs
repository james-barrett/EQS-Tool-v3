using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EQS_Tool
{
    public class CertificateProcessing
    {
        public static bool ProcessingCertificate(string baseDir, string remotePath, string department, string fileName, string file, string logFilePath, string timestamp)
        {
            string message = "";
            bool result = false;

            string destinationPath = Path.Combine(remotePath, fileName);
            string errorFilePath = Path.Combine(baseDir, department, "ERROR", fileName);

            // Test destination folders for duplicates
            if (!File.Exists(destinationPath))
            {
                try
                {
                    File.Copy(file, destinationPath);
                    message = $"Copied {file} to {destinationPath}";
                    Logger.Log(message, logFilePath, LogLevel.Info);
                    result = true;
                }
                catch (Exception ex)
                {
                    message = $"Failed to move {file} to ERROR folder. Exception: {ex.Message}";
                    Logger.Log(message, logFilePath, LogLevel.Info);
                }
            }
            else
            {
                message = $"Failed to copy {file} to {destinationPath}, file already exists.";
                Logger.Log(message, logFilePath, LogLevel.Info);
            }

            return result;
        }
    }
}
