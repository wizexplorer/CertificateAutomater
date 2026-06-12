using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Runtime.InteropServices;

namespace CertificateAutomater;

public static class CertificateAutomation
{
    private const string SearchText = "CERTIFICATE HOLDER";

    public static void Run(
        string templatePath,
        string certificateDirectory,
        bool convertToPdf,
        Action<string> log
    )
    {
        if (string.IsNullOrWhiteSpace(templatePath))
        {
            throw new Exception("You did not provide a template file path.");
        }

        if (!File.Exists(templatePath))
        {
            throw new Exception("The template file does not exist.");
        }

        if (!IsSupportedExcelFile(templatePath))
        {
            throw new Exception("Template file must be .xls or .xlsx.");
        }

        if (string.IsNullOrWhiteSpace(certificateDirectory))
        {
            throw new Exception("You did not provide a certificate folder path.");
        }

        if (!Directory.Exists(certificateDirectory))
        {
            throw new Exception("The certificate folder does not exist.");
        }

        string? pdfOutputDirectory = null;

        if (convertToPdf)
        {
            pdfOutputDirectory = PreparePdfOutputDirectory(templatePath);
            log($"PDF output folder: {pdfOutputDirectory}");
        }
        else
        {
            log("PDF generation is disabled.");
        }

        string[] certificateFiles = Directory
            .GetFiles(certificateDirectory)
            .Where(IsValidCertificateFileToProcess)
            .ToArray();

        if (certificateFiles.Length == 0)
        {
            log("No valid certificate files were found.");
            return;
        }

        log("");
        log($"Found {certificateFiles.Length} certificate file(s) to process.");
        log("");

        int successCount = 0;
        int failureCount = 0;

        foreach (string certificatePath in certificateFiles)
        {
            log("--------------------------------------------------");
            log($"Processing: {Path.GetFileName(certificatePath)}");

            try
            {
                ProcessCertificateFile(
                    templatePath,
                    certificatePath,
                    pdfOutputDirectory,
                    convertToPdf,
                    log
                );

                successCount++;
                log("Status: SUCCESS");
            }
            catch (Exception ex)
            {
                failureCount++;
                log("Status: FAILED");
                log($"Reason: {ex.Message}");
            }

            log("");
        }

        log("==================================================");
        log("Batch complete.");
        log($"Successful: {successCount}");
        log($"Failed: {failureCount}");
        log("==================================================");
    }

    // =========================================================
    // Main processing function
    // =========================================================

    private static void ProcessCertificateFile(
        string templatePath,
        string certificatePath,
        string? pdfOutputDirectory,
        bool convertToPdf,
        Action<string> log
    )
    {
        IWorkbook certificateWorkbook = LoadWorkbook(certificatePath);
        ISheet certificateSheet = certificateWorkbook.GetSheetAt(0);

        CellRangeAddress? certificateHeaderRange = FindBoldTextRange(
            certificateWorkbook,
            certificateSheet,
            SearchText
        );

        if (certificateHeaderRange == null)
        {
            throw new Exception($"Could not find bold '{SearchText}' in the first sheet of the certificate file.");
        }

        log($"Found certificate header at: {certificateSheet.SheetName}!{ToRangeAddress(certificateHeaderRange)}");

        int sourceStartRow = certificateHeaderRange.LastRow + 1;
        int sourceEndRow = certificateSheet.LastRowNum;

        int sourceStartColumn = certificateHeaderRange.FirstColumn;
        int sourceEndColumn = certificateHeaderRange.LastColumn;

        List<List<CellCopy>> copiedRows = CopyCellBlock(
            certificateSheet,
            sourceStartRow,
            sourceEndRow,
            sourceStartColumn,
            sourceEndColumn
        );

        log(
            $"Copied rows {sourceStartRow + 1} to {sourceEndRow + 1}, " +
            $"columns {ToColumnName(sourceStartColumn)} to {ToColumnName(sourceEndColumn)}."
        );

        string templateDirectory = Path.GetDirectoryName(templatePath)!;
        string templateExtension = Path.GetExtension(templatePath);
        string certificateFileNameWithoutExtension = Path.GetFileNameWithoutExtension(certificatePath);

        string outputFileName = certificateFileNameWithoutExtension + templateExtension;
        string outputPath = Path.Combine(templateDirectory, outputFileName);

        if (File.Exists(outputPath))
        {
            File.Copy(templatePath, outputPath, overwrite: true);
            log($"Existing output file was found and replaced: {outputPath}");
        }
        else
        {
            File.Copy(templatePath, outputPath);
            log($"Created new file from template: {outputPath}");
        }

        IWorkbook outputWorkbook = LoadWorkbook(outputPath);
        ISheet outputSheet = outputWorkbook.GetSheetAt(0);

        CellRangeAddress? templateHeaderRange = FindBoldTextRange(
            outputWorkbook,
            outputSheet,
            SearchText
        );

        if (templateHeaderRange == null)
        {
            throw new Exception($"Could not find bold '{SearchText}' in the copied template file.");
        }

        log($"Found template header at: {outputSheet.SheetName}!{ToRangeAddress(templateHeaderRange)}");

        int targetStartRow = templateHeaderRange.LastRow + 1;
        int targetStartColumn = templateHeaderRange.FirstColumn;

        int sourceColumnCount = sourceEndColumn - sourceStartColumn + 1;
        int targetColumnCount = templateHeaderRange.LastColumn - templateHeaderRange.FirstColumn + 1;

        if (sourceColumnCount != targetColumnCount)
        {
            log("Warning: source and target certificate-holder column ranges are not the same width.");
            log($"Source width: {sourceColumnCount}");
            log($"Target width: {targetColumnCount}");
            log("The program will paste using the smaller of the two widths.");
        }

        PasteCellBlock(
            outputSheet,
            copiedRows,
            targetStartRow,
            targetStartColumn,
            Math.Min(sourceColumnCount, targetColumnCount)
        );

        SaveWorkbook(outputWorkbook, outputPath);

        log($"Saved output Excel file: {outputPath}");

        if (convertToPdf)
        {
            if (string.IsNullOrWhiteSpace(pdfOutputDirectory))
            {
                throw new Exception("PDF output directory was not prepared.");
            }

            string pdfFileName = Path.GetFileNameWithoutExtension(outputPath) + ".pdf";
            string pdfPath = Path.Combine(pdfOutputDirectory, pdfFileName);

            ExportExcelToPdf(outputPath, pdfPath);

            log($"Saved PDF file: {pdfPath}");
        }
        else
        {
            log("Skipped PDF generation.");
        }
    }

    // =========================================================
    // PDF folder helper
    // =========================================================

    private static string PreparePdfOutputDirectory(string templatePath)
    {
        string templateDirectory = Path.GetDirectoryName(templatePath)!;

        string pdfDirectory = Path.Combine(templateDirectory, "PDF");
        // string pdfCertDirectory = Path.Combine(templateDirectory, "PDF-CERT");
        Directory.CreateDirectory(pdfDirectory);
        return pdfDirectory;

        // if (!Directory.Exists(pdfDirectory))
        // {
        //     Directory.CreateDirectory(pdfDirectory);
        //     return pdfDirectory;
        // }

        // if (Directory.Exists(pdfCertDirectory))
        // {
        //     Directory.Delete(pdfCertDirectory, recursive: true);
        // }

        // Directory.CreateDirectory(pdfCertDirectory);
        // return pdfCertDirectory;
    }

    // =========================================================
    // PDF export helper
    // =========================================================

    private static void ExportExcelToPdf(string excelPath, string pdfPath)
    {
        object? excelApp = null;
        object? workbooks = null;
        object? workbook = null;

        try
        {
            if (File.Exists(pdfPath))
            {
                File.Delete(pdfPath);
            }

            Type? excelType = Type.GetTypeFromProgID("Excel.Application");

            if (excelType == null)
            {
                throw new Exception("Microsoft Excel COM automation is not available. Excel may not be installed or registered correctly.");
            }

            excelApp = Activator.CreateInstance(excelType);

            if (excelApp == null)
            {
                throw new Exception("Could not create an Excel application instance.");
            }

            dynamic excel = excelApp;

            excel.Visible = false;
            excel.DisplayAlerts = false;

            workbooks = excel.Workbooks;
            dynamic dynamicWorkbooks = workbooks;

            workbook = dynamicWorkbooks.Open(excelPath, ReadOnly: true);
            dynamic dynamicWorkbook = workbook;

            // 0 = PDF
            dynamicWorkbook.ExportAsFixedFormat(
                0,
                pdfPath
            );

            dynamicWorkbook.Close(false);
            excel.Quit();
        }
        finally
        {
            if (workbook != null)
            {
                Marshal.ReleaseComObject(workbook);
            }

            if (workbooks != null)
            {
                Marshal.ReleaseComObject(workbooks);
            }

            if (excelApp != null)
            {
                Marshal.ReleaseComObject(excelApp);
            }

            workbook = null;
            workbooks = null;
            excelApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    // =========================================================
    // File filtering helpers
    // =========================================================

    private static bool IsValidCertificateFileToProcess(string path)
    {
        if (!IsSupportedExcelFile(path))
        {
            return false;
        }

        string fileNameWithoutExtension = Path
            .GetFileNameWithoutExtension(path)
            .ToLowerInvariant();

        if (fileNameWithoutExtension.Contains("template"))
        {
            return false;
        }

        if (fileNameWithoutExtension.Contains("000"))
        {
            return false;
        }

        string fileName = Path.GetFileName(path);

        if (fileName.StartsWith("~$"))
        {
            return false;
        }

        return true;
    }

    private static bool IsSupportedExcelFile(string path)
    {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        return extension == ".xls" || extension == ".xlsx";
    }

    // =========================================================
    // Workbook helpers
    // =========================================================

    private static IWorkbook LoadWorkbook(string path)
    {
        string extension = Path.GetExtension(path).ToLowerInvariant();

        using FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);

        if (extension == ".xls")
        {
            return new HSSFWorkbook(fileStream);
        }

        if (extension == ".xlsx")
        {
            return new XSSFWorkbook(fileStream);
        }

        throw new ArgumentException("Unsupported Excel file type.");
    }

    private static void SaveWorkbook(IWorkbook workbook, string path)
    {
        using FileStream outputStream = new FileStream(path, FileMode.Create, FileAccess.Write);
        workbook.Write(outputStream);
    }

    // =========================================================
    // Search helpers
    // =========================================================

    private static CellRangeAddress? FindBoldTextRange(
        IWorkbook workbook,
        ISheet sheet,
        string searchText
    )
    {
        for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            IRow? row = sheet.GetRow(rowIndex);

            if (row == null)
            {
                continue;
            }

            if (row.FirstCellNum < 0 || row.LastCellNum < 0)
            {
                continue;
            }

            for (int cellIndex = row.FirstCellNum; cellIndex < row.LastCellNum; cellIndex++)
            {
                ICell? cell = row.GetCell(cellIndex);

                if (cell == null)
                {
                    continue;
                }

                string cellText = GetCellText(cell).Trim();

                bool textMatches = cellText.Equals(
                    searchText,
                    StringComparison.OrdinalIgnoreCase
                );

                if (!textMatches)
                {
                    continue;
                }

                bool isBold = IsCellBold(workbook, cell);

                if (!isBold)
                {
                    continue;
                }

                CellRangeAddress? mergedRange = FindMergedRangeContainingCell(
                    sheet,
                    rowIndex,
                    cellIndex
                );

                if (mergedRange != null)
                {
                    return mergedRange;
                }

                return new CellRangeAddress(rowIndex, rowIndex, cellIndex, cellIndex);
            }
        }

        return null;
    }

    private static string GetCellText(ICell cell)
    {
        if (cell.CellType == CellType.String)
        {
            return cell.StringCellValue ?? "";
        }

        if (cell.CellType == CellType.Numeric)
        {
            return cell.NumericCellValue.ToString();
        }

        if (cell.CellType == CellType.Boolean)
        {
            return cell.BooleanCellValue.ToString();
        }

        if (cell.CellType == CellType.Formula)
        {
            return GetFormulaCellText(cell);
        }

        return "";
    }

    private static string GetFormulaCellText(ICell cell)
    {
        if (cell.CachedFormulaResultType == CellType.String)
        {
            return cell.StringCellValue ?? "";
        }

        if (cell.CachedFormulaResultType == CellType.Numeric)
        {
            return cell.NumericCellValue.ToString();
        }

        if (cell.CachedFormulaResultType == CellType.Boolean)
        {
            return cell.BooleanCellValue.ToString();
        }

        return "";
    }

    private static bool IsCellBold(IWorkbook workbook, ICell cell)
    {
        ICellStyle cellStyle = cell.CellStyle;
        IFont font = workbook.GetFontAt(cellStyle.FontIndex);

        return font.IsBold;
    }

    // =========================================================
    // Copy/paste helpers
    // =========================================================

    private static List<List<CellCopy>> CopyCellBlock(
        ISheet sheet,
        int startRow,
        int endRow,
        int startColumn,
        int endColumn
    )
    {
        List<List<CellCopy>> copiedRows = new List<List<CellCopy>>();

        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            IRow? row = sheet.GetRow(rowIndex);
            List<CellCopy> copiedCellsInRow = new List<CellCopy>();

            for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
            {
                ICell? cell = row?.GetCell(columnIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);

                if (cell == null)
                {
                    copiedCellsInRow.Add(CellCopy.Blank());
                }
                else
                {
                    copiedCellsInRow.Add(ReadCell(cell));
                }
            }

            copiedRows.Add(copiedCellsInRow);
        }

        return copiedRows;
    }

    private static void PasteCellBlock(
        ISheet sheet,
        List<List<CellCopy>> copiedRows,
        int targetStartRow,
        int targetStartColumn,
        int maxColumnCount
    )
    {
        for (int rowOffset = 0; rowOffset < copiedRows.Count; rowOffset++)
        {
            int targetRowIndex = targetStartRow + rowOffset;

            IRow targetRow = sheet.GetRow(targetRowIndex) ?? sheet.CreateRow(targetRowIndex);

            List<CellCopy> copiedCellsInRow = copiedRows[rowOffset];

            for (int columnOffset = 0; columnOffset < maxColumnCount; columnOffset++)
            {
                int targetColumnIndex = targetStartColumn + columnOffset;

                if (IsInsideMergedRegionButNotTopLeft(sheet, targetRowIndex, targetColumnIndex))
                {
                    continue;
                }

                ICell targetCell = targetRow.GetCell(targetColumnIndex)
                    ?? targetRow.CreateCell(targetColumnIndex);

                CellCopy sourceCell = copiedCellsInRow[columnOffset];

                WriteCell(targetCell, sourceCell);
            }
        }
    }

    private static CellCopy ReadCell(ICell cell)
    {
        if (cell.CellType == CellType.String)
        {
            return CellCopy.Text(cell.StringCellValue ?? "");
        }

        if (cell.CellType == CellType.Numeric)
        {
            return CellCopy.Number(cell.NumericCellValue);
        }

        if (cell.CellType == CellType.Boolean)
        {
            return CellCopy.Boolean(cell.BooleanCellValue);
        }

        if (cell.CellType == CellType.Formula)
        {
            return ReadFormulaCachedValue(cell);
        }

        return CellCopy.Blank();
    }

    private static CellCopy ReadFormulaCachedValue(ICell cell)
    {
        if (cell.CachedFormulaResultType == CellType.String)
        {
            return CellCopy.Text(cell.StringCellValue ?? "");
        }

        if (cell.CachedFormulaResultType == CellType.Numeric)
        {
            return CellCopy.Number(cell.NumericCellValue);
        }

        if (cell.CachedFormulaResultType == CellType.Boolean)
        {
            return CellCopy.Boolean(cell.BooleanCellValue);
        }

        return CellCopy.Blank();
    }

    private static void WriteCell(ICell targetCell, CellCopy sourceCell)
    {
        if (sourceCell.Kind == CellCopyKind.Blank)
        {
            targetCell.SetBlank();
            return;
        }

        if (sourceCell.Kind == CellCopyKind.Text)
        {
            targetCell.SetCellValue(sourceCell.TextValue ?? "");
            return;
        }

        if (sourceCell.Kind == CellCopyKind.Number)
        {
            targetCell.SetCellValue(sourceCell.NumberValue);
            return;
        }

        if (sourceCell.Kind == CellCopyKind.Boolean)
        {
            targetCell.SetCellValue(sourceCell.BooleanValue);
            return;
        }
    }

    // =========================================================
    // Merged-cell helpers
    // =========================================================

    private static CellRangeAddress? FindMergedRangeContainingCell(
        ISheet sheet,
        int rowIndex,
        int columnIndex
    )
    {
        for (int i = 0; i < sheet.NumMergedRegions; i++)
        {
            CellRangeAddress range = sheet.GetMergedRegion(i);

            bool rowIsInside =
                rowIndex >= range.FirstRow &&
                rowIndex <= range.LastRow;

            bool columnIsInside =
                columnIndex >= range.FirstColumn &&
                columnIndex <= range.LastColumn;

            if (rowIsInside && columnIsInside)
            {
                return range;
            }
        }

        return null;
    }

    private static bool IsInsideMergedRegionButNotTopLeft(
        ISheet sheet,
        int rowIndex,
        int columnIndex
    )
    {
        CellRangeAddress? range = FindMergedRangeContainingCell(
            sheet,
            rowIndex,
            columnIndex
        );

        if (range == null)
        {
            return false;
        }

        bool isTopLeft =
            rowIndex == range.FirstRow &&
            columnIndex == range.FirstColumn;

        return !isTopLeft;
    }

    // =========================================================
    // Address helpers
    // =========================================================

    private static string ToRangeAddress(CellRangeAddress range)
    {
        string start = ToCellAddress(range.FirstRow, range.FirstColumn);
        string end = ToCellAddress(range.LastRow, range.LastColumn);

        if (start == end)
        {
            return start;
        }

        return $"{start}:{end}";
    }

    private static string ToCellAddress(int rowIndex, int columnIndex)
    {
        string columnName = ToColumnName(columnIndex);
        int excelRowNumber = rowIndex + 1;

        return $"{columnName}{excelRowNumber}";
    }

    private static string ToColumnName(int columnIndex)
    {
        string columnName = "";
        int dividend = columnIndex + 1;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    // =========================================================
    // Helper types
    // =========================================================

    private enum CellCopyKind
    {
        Blank,
        Text,
        Number,
        Boolean
    }

    private class CellCopy
    {
        public CellCopyKind Kind { get; private set; }
        public string? TextValue { get; private set; }
        public double NumberValue { get; private set; }
        public bool BooleanValue { get; private set; }

        public static CellCopy Blank()
        {
            return new CellCopy
            {
                Kind = CellCopyKind.Blank
            };
        }

        public static CellCopy Text(string value)
        {
            return new CellCopy
            {
                Kind = CellCopyKind.Text,
                TextValue = value
            };
        }

        public static CellCopy Number(double value)
        {
            return new CellCopy
            {
                Kind = CellCopyKind.Number,
                NumberValue = value
            };
        }

        public static CellCopy Boolean(bool value)
        {
            return new CellCopy
            {
                Kind = CellCopyKind.Boolean,
                BooleanValue = value
            };
        }
    }
}