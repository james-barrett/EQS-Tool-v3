using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using System;
using System.Data;
using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Page = UglyToad.PdfPig.Content.Page;


namespace EQS_Tool
{
    class Program
    {
        static void Main(string[] args)
        {
            // Toggle bool to switch to production directory
            bool isProduction = false;

            string pathCur = Directory.GetCurrentDirectory();
            string basePath;

            if (isProduction)
            {
                basePath = pathCur;
            }
            else
            {
                basePath = pathCur.Split(new[] { "\\bin" }, StringSplitOptions.None)[0];
            }

            // Set current timestamp for unique file creation
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmm");

            // Create log dir if non exists
            // Create log file using timestamp
            string[] dirs = [
                Path.Combine(basePath, "_logs"),
                Path.Combine(basePath, "Certificates"),
                Path.Combine(basePath, "Audit Data")
            ];

            string logFilePath = Path.Combine(dirs[0], $"Log-{timestamp}.txt");

            foreach (string dir in dirs)
            {
                var result = DirectoryFunctions.Create(dir);
                LogLevel logLevel = result.success ? LogLevel.Info : LogLevel.Info;
                Logger.Log(result.message, logFilePath, logLevel);
            }

            

            // Load all values from config file or exit.
            ApplicationConfig appConfig = new ApplicationConfig();

            try
            {
                appConfig = ApplicationConfig.Load();
            }
            catch (ConfigurationException ex)
            {
                Logger.Log($"Error processing configuration file: {ex.Message}", logFilePath, LogLevel.Error);
                Helpers.CleanExit();
            }

            // manager function folders - will likely be expanded in future
            string[] auditFolders = [
                Path.Combine(dirs[2], "Evotix")
            ];

            // if not certificate mode then we are running in manager mode
            if (!appConfig.Settings.CertificateMode)
            {
                Logger.Log($"Running in audit mode.", logFilePath, LogLevel.Info);

                foreach (string auditFolder in auditFolders)
                {
                    var result = DirectoryFunctions.Create(auditFolder);
                    LogLevel logLevel = result.success ? LogLevel.Info : LogLevel.Info;
                    Logger.Log(result.message, logFilePath, logLevel);
                }

                // months as strings for file naming
                List<string> monthsAsStrings = new List<string>();
                monthsAsStrings = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

                // process only csv files in root of directory
                foreach (string auditFolder in auditFolders)
                {
                    string[] auditDate = GenerateAuditData.InputAuditDate();

                    // file path name built from the user input month/year
                    string auditExcelFilePath = Path.Combine(auditFolder, Path.Combine(auditFolder, $"Audit-Data-{monthsAsStrings[int.Parse(auditDate[1]) - 1]}-{auditDate[0]}-{timestamp}.xlsx"));
                    string auditFolderPath = Path.Combine(auditFolder, auditFolder);
                    string[] files = Directory.GetFiles(auditFolderPath, "*.csv", SearchOption.TopDirectoryOnly);
                    Logger.Log($"Processing {auditFolder} found {files.Length}.", logFilePath, LogLevel.Info);

                    foreach (string file in files)
                    {                        
                        Logger.Log($"Processing {file}.", logFilePath, LogLevel.Info);
                        GenerateAuditData.ProcessAuditData(file, auditDate, auditExcelFilePath, appConfig.Settings.EQSTeam, monthsAsStrings[int.Parse(auditDate[1]) - 1], auditDate[0]);
                    }
                }
                // if not certificate mode then code flow ends here
                return;
            }

            // we are in certificate mode as the above if check didnt trigger
            string excelAddressListPath = Path.Combine(basePath, "PROPERTIES.xlsx");
            DataTable addressTable = DataTableFunctions.LoadDataTable(excelAddressListPath, logFilePath, timestamp);

            // department sub folders, needed for logic
            string[] departments = ["EH", "RR"];
            string[] folders = ["ERROR", "UPRN_ERROR", "ADDRESS_CHECK_FAILED", "PROCESSED"];

            string certificateDirectory = dirs[1];

            // check and create if missing or exit
            foreach (string dept in departments)
            {
                foreach (string folder in folders)
                {
                    string folderPath = Path.Combine(certificateDirectory, dept, folder);

                    if (Directory.Exists(folderPath)) continue;

                    try
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    catch (Exception ex)
                    {
                        throw new IOException($"Failed to create directory: {folderPath}", ex);
                    }
                }
            }

            // process only pdf files in root of directory per dept sub folder
            foreach (string department in departments)
            {
                string deptFolderPath = Path.Combine(certificateDirectory, department);
                string[] files = Directory.GetFiles(deptFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);

                Logger.Log($"Processing {department} found {files.Length}.", logFilePath, LogLevel.Info);

                foreach (string file in files)
                {
                    Thread.Sleep(2000);
                    Logger.Log($"-------------------------------------------", logFilePath, LogLevel.Info);
                    Logger.Log($"Processing {file}.", logFilePath, LogLevel.Info);

                    PdfDocument doc;
                    IEnumerable<Word> words;
                    Page[] pages;

                    // try to load all pages of pdf into array for processing
                    // on error move to ERROR folder
                    try
                    {
                        doc = PdfDocument.Open(file);
                        int pageCount = doc.GetPages().Count();
                        pages = new Page[pageCount + 1];

                        // use first member of array at index 1 to keep logic simple as pdf page index starts at 1
                        for (int i = 1; i < pageCount + 1; i++)
                        {
                            pages[i] = doc.GetPage(i);
                        }
                    }
                    catch (Exception ex)
                    {
                        // error path used multiple times so set as string variable
                        string errorFilePath = Path.Combine(deptFolderPath, "ERROR", Path.GetFileName(file));
                        Logger.Log($"Error occurred while trying to extract PDF data from {file}.", logFilePath, LogLevel.Error);

                        if (!File.Exists(errorFilePath))
                        {
                            try
                            {
                                File.Move(file, errorFilePath);
                                Logger.Log($"Moved {Path.GetFileName(file)} to ERROR folder due to exception: {ex.Message}", logFilePath, LogLevel.Info);
                            }
                            catch (Exception moveEx)
                            {
                                Logger.Log($"Failed to move file {Path.GetFileName(file)} to ERROR folder. Exception: {moveEx.Message}", logFilePath, LogLevel.Error);
                            }
                        }
                        else
                        {
                            Logger.Log($"Error file already exists: {errorFilePath}", logFilePath, LogLevel.Info);
                        }
                        continue;
                    }

                    // use first page of pdf which is at index 1 of array
                    // use this to get certificate type using helper function
                    Page page = pages[1];
                    words = page.GetWords();
                    string pageText = string.Join(" ", page.GetWords());
                    string certificateType = CertificateFunctions.CertificateType(pageText);

                    // if enabled various tools below
                    if (appConfig.Settings.CertificateDataCollection)
                    {
                        CertificateFunctions.CertificateDataCollection();
                    }

                    if (appConfig.Settings.PrintTextCoordinates)
                    {
                        DevTools.PrintCoordsToConsole(page);
                    }

                    if (appConfig.Settings.DrawRectsToPDF)
                    {
                        DevTools.DrawRectsToPDF(file, deptFolderPath, certificateType, appConfig, timestamp);
                    }

                    // if returned value from CertificateType function is ERROR we have an invalid file. 
                    // move to ERROR folder
                    if (certificateType == "ERROR")
                    {
                        // error path used multiple times so set as string variable
                        string errorFilePath = Path.Combine(deptFolderPath, "ERROR", Path.GetFileName(file));
                        Logger.Log($"Certificate error, {certificateType} is not a valid document.", logFilePath, LogLevel.Info);

                        if (!File.Exists(errorFilePath))
                        {
                            try
                            {
                                File.Move(file, errorFilePath);
                                Logger.Log($"Moved {Path.GetFileName(file)} to ERROR folder.", logFilePath, LogLevel.Info);
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"Failed to move file {file} to ERROR folder. Exception: {ex.Message}", logFilePath, LogLevel.Info);
                            }
                        }
                        else
                        {
                            Logger.Log($"Error file already exists: {errorFilePath}", logFilePath, LogLevel.Info);
                        }
                        continue;
                    }

                    // helper function CertificateData uses certificateType to locate correct Rect coords from config
                    // then parses pdf to retrieve data from those locations.
                    string[] certificateData = CertificateFunctions.CertificateData(certificateType, pages, appConfig);

                    // Set variables from certificate data
                    string jobRef = certificateData[0];
                    // certificateData[1];
                    string certificateNumber = certificateData[2];
                    string date = certificateData[3];
                    string addressLine1 = certificateData[4];
                    string addressLine2 = certificateData[5];
                    string postcode = certificateData[6];
                    string engineer = certificateData[7];
                    string supervisor = certificateData[8];
                    string? result = null;

                    string occupier = certificateData[10];

                    if (certificateType == "EICR")
                    {
                        result = CertificateFunctions.EICRResult(page, appConfig);

                        if (result == "ERROR")
                        {
                            // error path used multiple times so set as string variable
                            string errorFilePath = Path.Combine(deptFolderPath, "ERROR", Path.GetFileName(file));
                            Logger.Log($"EICR Result error, {result} is not a valid result.", logFilePath, LogLevel.Info);
                            if (!File.Exists(errorFilePath))
                            {
                                try
                                {
                                    File.Move(file, errorFilePath);
                                    Logger.Log($"Moved {Path.GetFileName(file)} to ERROR folder.", logFilePath, LogLevel.Info);
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log($"Failed to move file {file} to ERROR folder. Exception: {ex.Message}", logFilePath, LogLevel.Info);
                                }
                            }
                            else
                            {
                                Logger.Log($"Error file already exists: {errorFilePath}", logFilePath, LogLevel.Info);
                            }
                            continue;
                        }
                    }

                    if (appConfig.Settings.PrintDataToConsole)
                    {
                        foreach (string key in certificateData)
                        {
                            Console.WriteLine(key);
                        }
                    }

                    // if certificateType is DFHN we need to get UPRN using the address
                    // the address should be perfect from EQS checks so should be no issues
                    if (certificateType == "DFHN")
                    {
                        string address = $"{certificateData[4]} {certificateData[5]} {certificateData[6]}";
                        certificateData[1] = DataTableFunctions.GetUPRNFromDataTable(addressTable, address, logFilePath, timestamp);
                    }

                    // if certificateType is PARTP we need to get UPRN using the address
                    // the address is incomplete and does not match the certificate usually missing county
                    // using contains to match the first part of the address, will monitor for accuracy
                    if (certificateType == "PARTP")
                    {
                        certificateData[1] = DataTableFunctions.GetUPRNFromDataTable(addressTable, certificateData[4], logFilePath, timestamp);
                    }

                    // set UPRN, either from initial load of certificateData or amended via if DFHN or PARTP
                    string uprn = certificateData[1];

                    // format date if of type (DD-MM-YYYY - DD-MM-YYYY) testing was carried out over a span of time
                    // use end date for certificate naming
                    try
                    {
                        // check if the date is not null or empty before processing
                        if (!string.IsNullOrWhiteSpace(date))
                        {
                            if (date.Contains("-"))
                            {
                                string[] dateArray = date.Split('-');

                                if (dateArray.Length > 1 && DateTime.TryParse(dateArray[1].Trim(), out DateTime parsedDate))
                                {
                                    date = dateArray[1].Trim();
                                    certificateData[3] = date;
                                }
                            }
                            else
                            {
                                date = date.Trim();
                                certificateData[3] = date;
                            }
                        }
                        else
                        {
                            Logger.Log($"Error parsing date from {file}", logFilePath, LogLevel.Info);
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"Error parsing date from {file}, Exception: {ex}", logFilePath, LogLevel.Info);
                        continue;
                    }

                    // get file name using helper function GenerateFileName
                    string fileName = FileNameHelpers.GenerateFileName(certificateType, certificateData, appConfig.Settings.NamingFormat);

                    // get address from DB based on UPRN
                    // if UPRN_ERROR returned from function, move to UPRN_ERROR folder
                    string addressDB = DataTableFunctions.GetAddressFromDataTable(addressTable, uprn.ToUpper(), logFilePath, timestamp);

                    if (addressDB == "UPRN_ERROR")
                    {
                        // error path used multiple times so set as string variable
                        string errorFilePath = Path.Combine(deptFolderPath, "UPRN_ERROR", Path.GetFileName(file));

                        if (!File.Exists(errorFilePath))
                        {
                            try
                            {
                                File.Move(file, errorFilePath);
                                Logger.Log($"UPRN Error, moving {file} to UPRN_ERROR folder.", logFilePath, LogLevel.Info);
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"Failed to move {file} to UPRN_ERROR folder. Exception: {ex.Message}", logFilePath, LogLevel.Info);
                            }
                        }
                        else
                        {
                            Logger.Log($"Error moving {Path.GetFileName(file)} to UPRN_ERROR folder, file already exists.", logFilePath, LogLevel.Info);
                        }
                        continue;
                    }

                    // if not DFHN or PARTP as we dont need to check those for valid addresses with no UPRN to reference
                    // if address errors likley will not match with correct UPRN from DataTable funtions and be unable to process
                    if (certificateType != "DFHN" && certificateType != "PARTP")
                    {
                        // use helper function AddressFuzzyMatchScore to validate address
                        float[] scores = StringFunctions.AddressFuzzyMatchScore(addressDB, addressLine1, addressLine2, postcode);

                        // process address check failure or continue
                        if (scores[0] != 0 || scores[1] != 0 || scores[2] >= appConfig.Settings.ScoreThreshold)
                        {
                            // error path used multiple times so set as string variable 
                            string failedFilePath = Path.Combine(deptFolderPath, "ADDRESS_CHECK_FAILED", Path.GetFileName(file));

                            if (!File.Exists(failedFilePath))
                            {
                                try
                                {
                                    File.Move(file, failedFilePath);
                                    Logger.Log($"Address check failed.", logFilePath, LogLevel.Info);
                                    Logger.Log($"Expected: {addressDB}", logFilePath, LogLevel.Info);
                                    Logger.Log($"Received: {addressLine1} {addressLine2} {postcode}", logFilePath, LogLevel.Info);
                                    Logger.Log($"Moving {file} to ADDRESS_CHECK_FAILED folder.", logFilePath, LogLevel.Info);
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log($"Failed to move {file} to ADDRESS_CHECK_FAILED folder. Exception: {ex.Message}", logFilePath, LogLevel.Info);
                                }
                            }
                            else
                            {
                                Logger.Log($"Error moving {Path.GetFileName(file)} to ADDRESS_CHECK_FAILED folder, file already exists.", logFilePath, LogLevel.Info);
                            }
                            continue;
                        }
                    }

                    // if auto move files
                    if (appConfig.Settings.AutoMoveFiles)
                    {
                        // check to make sure we arnt processing Empty Homes files from other department folders
                        // check if not PARTP as does not contain occupier details
                        if (certificateType != "PARTP")
                        {
                            if (department != "EH" && (certificateData[10].ToUpper().Contains("EMPTY") || certificateData[10].ToUpper().Contains("VOID")))
                            {
                                Logger.Log($"Error possible Empty Homes certificate found in incorrect folder, skipping...", logFilePath, LogLevel.Info);
                                continue;
                            }
                            else if (certificateData[2].ToUpper().Contains("DRAFT"))
                            {
                                Logger.Log($"Error Draft certificate found, skipping...", logFilePath, LogLevel.Info);
                                continue;
                            }
                            else if (certificateData[2] == null)
                            {
                                Logger.Log($"Error certificate number not found, skipping...", logFilePath, LogLevel.Info);
                                continue;
                            }
                        }

                        // Use bools to check for successful operations before we delete local file
                        bool status;

                        if (certificateType == "EICR" && result == "UNSAT")
                        {
                            

                            status = CertificateProcessing.ProcessingCertificate(
                                certificateDirectory, appConfig.Paths.TGPUNSATFilePath, department, fileName, file, logFilePath, timestamp);

                            if (status)
                            {
                                try
                                {
                                    File.Delete(file);
                                    Logger.Log($"Local file {file} Deleted.", logFilePath, LogLevel.Info);
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log($"File actions not completed. Exception: {ex}.", logFilePath, LogLevel.Info);
                                }
                            }
                            else
                            {
                                Logger.Log($"Error, processing certificate. Its possible the shared drive is not available", logFilePath, LogLevel.Info);
                            }

                            continue;
                        }

                        // Process based on department
                        switch (department)
                        {
                            case "EH":

                                if (certificateType == "EICR")
                                {
                                    status = CertificateProcessing.ProcessingCertificate(
                                        certificateDirectory, appConfig.Paths.TGPEHFilePath, department, fileName, file, logFilePath, timestamp);

                                    if (status)
                                    {
                                        if (appConfig.Settings.AutoEmailEHEICR)
                                        {
                                            Email.Send(appConfig.Settings.EHEmailAddress, file, fileName, addressLine1 + " " + addressLine2 + " " + postcode);
                                        }

                                        try
                                        {
                                            File.Delete(file);
                                            Logger.Log($"Local file {file} Deleted.", logFilePath, LogLevel.Info);
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.Log($"File actions not completed. Exception: {ex}.", logFilePath, LogLevel.Info);
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log($"Error, processing certificate. Its possible the shared drive is not available", logFilePath, LogLevel.Info);
                                    }
                                }
                                else
                                {
                                    if (CertificateProcessing.ProcessingCertificate(certificateDirectory, appConfig.Paths.TGPEHFilePath, department, fileName, file, logFilePath, timestamp))
                                    {
                                        if (appConfig.Settings.AutoEmailEHOther)
                                        {
                                            Email.Send(appConfig.Settings.EHEmailAddress, file, fileName, addressLine1 + " " + addressLine2 + " " + postcode);
                                        }

                                        try
                                        {
                                            File.Delete(file);
                                            Logger.Log($"Local file {file} Deleted.", logFilePath, LogLevel.Info);
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.Log($"File actions not completed. Exception: {ex}.", logFilePath, LogLevel.Info);
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log($"Error, processing certificate. Its possible the shared drive is not available", logFilePath, LogLevel.Info);
                                    }
                                }
                                continue;

                            
                            case "RR":

                                if (certificateType == "EICR")
                                {                                    

                                    status = CertificateProcessing.ProcessingCertificate(
                                        certificateDirectory, appConfig.Paths.TGPRRFilePath, department, fileName, file, logFilePath, timestamp);

                                    if (status)
                                    {
                                        if (appConfig.Settings.AutoEmailRREICR)
                                        {
                                            Email.Send(appConfig.Settings.RREmailAddress, file, fileName, addressLine1 + " " + addressLine2 + " " + postcode);
                                        }

                                        try
                                        {
                                            File.Delete(file);
                                            Logger.Log($"Local file {file} Deleted.", logFilePath, LogLevel.Info);
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.Log($"File actions not completed. Exception: {ex}.", logFilePath, LogLevel.Info);
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log($"Error, processing certificate. Its possible the shared drive is not available", logFilePath, LogLevel.Info);
                                    }
                                }
                                else
                                {
                                    if (CertificateProcessing.ProcessingCertificate(certificateDirectory, appConfig.Paths.TGPRRFilePath, department, fileName, file, logFilePath, timestamp))
                                    {
                                        if (appConfig.Settings.AutoEmailRROther)
                                        {
                                            Email.Send(appConfig.Settings.RREmailAddress, file, fileName, addressLine1 + " " + addressLine2 + " " + postcode);
                                        }

                                        try
                                        {
                                            File.Delete(file);
                                            Logger.Log($"Local file {file} Deleted.", logFilePath, LogLevel.Info);
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.Log($"File actions not completed. Exception: {ex}.", logFilePath, LogLevel.Info);
                                        }
                                    }
                                    else
                                    {
                                        Logger.Log($"Error, processing certificate. Its possible the shared drive is not available", logFilePath, LogLevel.Info);
                                    }
                                }
                                continue;


                            default:
                                break;
                        }
                    }
                    else
                    {
                        string processedFilePath = Path.Combine(deptFolderPath, "PROCESSED", fileName);

                        if (!File.Exists(processedFilePath))
                        {
                            try
                            {
                                File.Copy(file, processedFilePath);
                                File.Delete(file);
                                Logger.Log($"File {fileName} moved to PROCESSED directory.", logFilePath, LogLevel.Info);

                                
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"Error processing {file}. Exception: {ex.Message}", logFilePath, LogLevel.Info);
                            }
                        }
                        continue;
                    }
                }               
            }
        }
    }
}