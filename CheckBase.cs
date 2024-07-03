using System.Collections.Generic;
using System.Linq;
using Models.Forms;
using OfficeOpenXml;
using System.Globalization;
using System.IO;
using System;
using System.ComponentModel;

namespace Client_App.Commands.SyncCommands.CheckForm;

public abstract class CheckBase
{
    private protected static List<Dictionary<string, string>> OKSM = new();

    private protected static List<Dictionary<string, string>> R = new();

    private protected static readonly bool DB_Ignore = true;

    private protected static readonly bool ZRI_Ignore = true;

    private protected static readonly bool MZA_Ignore = true;

    private protected static readonly Regex OkpoRegex = new(@"^\d{8}([0123456789_]\d{5})?$");

    #region CheckRepPeriod

    private protected static List<CheckError> CheckRepPeriod(List<Form1> forms, Report rep)
    {
        List<CheckError> result = new();
        if (!DateOnly.TryParse(rep.EndPeriod_DB, out var endPeriod)) return result;
        var minOpDate = DateOnly.MinValue;
        var opCode = string.Empty;
        var line = 0;
        string[] operationCodeWithDeadline1 = { "71" };
        string[] operationCodeWithDeadline5 = { "73", "74", "75" };
        string[] operationCodeWithDeadline10 =
        {
            "11", "12", "15", "17", "18", "21", "22", "25", "27", "28", "29", "31", "32", "35", "37", "38", "39",
            "41", "42", "43", "46", "47", "48", "53", "54", "58", "61", "62", "63", "64", "65", "66", "67", "68",
            "72", "81", "82", "83", "84", "85", "86", "87", "88", "97", "98", "99"
        };
        var formNum = rep.FormNum_DB.Replace(".", "");
        foreach (var form in forms)
        {
            var curOpCode = form.OperationCode_DB ?? string.Empty;
            var opDatePlus = DateOnly.MinValue;
            if (!DateOnly.TryParse(form.OperationDate_DB, out var opDate)) continue;
            if (operationCodeWithDeadline1.Contains(curOpCode)) opDatePlus = opDate.AddDays(1);
            if (operationCodeWithDeadline5.Contains(curOpCode)) opDatePlus = opDate.AddDays(5);
            if (operationCodeWithDeadline10.Contains(curOpCode)) opDatePlus = opDate.AddDays(10);
            if (opDatePlus > minOpDate)
            {
                minOpDate = opDate;
                opCode = form.OperationCode_DB ?? string.Empty;
                line = form.NumberInOrder_DB;
            }
        }
        if (operationCodeWithDeadline10.Contains(opCode) && WorkdaysBetweenDates(minOpDate, endPeriod) > 10)
        {
            result.Add(new CheckError
            {
                FormNum = $"form_{formNum}",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = forms[line].OperationDate_DB,
                Message = $"Дата операции {minOpDate} превышает дату окончания отчетного периода {rep.EndPeriod_DB} " +
                          $"более чем на 10 рабочих дней."
            });
        }
        else if (operationCodeWithDeadline5.Contains(opCode) && WorkdaysBetweenDates(minOpDate, endPeriod) > 5)
        {
            result.Add(new CheckError
            {
                FormNum = $"form_{formNum}",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = forms[line].OperationDate_DB,
                Message = $"Дата операции {minOpDate} превышает дату окончания отчетного периода {rep.EndPeriod_DB} " +
                          $"более чем на 5 рабочих дней."
            });
        }
        else if (operationCodeWithDeadline1.Contains(opCode) && WorkdaysBetweenDates(minOpDate, endPeriod) > 1)
        {
            result.Add(new CheckError
            {
                FormNum = $"form_{formNum}",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = forms[line].OperationDate_DB,
                Message = $"Дата операции {minOpDate} превышает дату окончания отчетного периода {rep.EndPeriod_DB} " +
                          $"более чем на 1 рабочий день."
            });
        }

        return result;
    }

    #endregion

    #region CheckNotePresence

    private protected static bool CheckNotePresence(List<Note> notes, int line, byte graphNumber)
    {
        var valid = false;
        foreach (var note in notes)
        {
            if (note.RowNumber_DB == null) continue;
            var noteRowsReal = note.RowNumber_DB.Replace(" ", string.Empty);
            List<int> noteRowsFinalInt = new();
            List<string> noteRowsRealStr = new(noteRowsReal.Split(','));
            foreach (var noteRowCluster in noteRowsRealStr)
            {
                if (noteRowCluster.Contains('-'))
                {
                    var noteRowBounds = noteRowCluster.Split('-');
                    if (noteRowBounds.Length != 2
                        || !int.TryParse(noteRowBounds[0], out var noteRowsBegin)
                        || !int.TryParse(noteRowBounds[1], out var noteRowsEnd)) continue;
                    for (var i = noteRowsBegin; i <= noteRowsEnd; i++)
                    {
                        noteRowsFinalInt.Add(i);
                    }
                }
                else
                {
                    if (int.TryParse(noteRowCluster, out var noteRowClusterInt))
                    {
                        noteRowsFinalInt.Add(noteRowClusterInt);
                    }
                }
            }
            if (noteRowsFinalInt.Any(noteRowNumber =>
                    noteRowNumber == line + 1
                    && note.GraphNumber_DB == graphNumber.ToString()
                    && !string.IsNullOrWhiteSpace(note.Comment_DB)))
            {
                valid = true;
                break;
            }
        }
        return valid;
    }

    #endregion

    #region ConvertStringToExponential

    protected static string ConvertStringToExponential(string? str)
    {
        str ??= "";
        return str
            .ToLower()
            .Trim()
            .TrimStart('(')
            .TrimEnd(')')
            .Replace(".", ",")
            .Replace('е', 'e');
    }

    #endregion

    #region CustomComparator

    private protected class CustomNullStringWithTrimComparer : IComparer<string>
    {
        public int Compare(string? x, string? y)
        {
            var strA = (x ?? string.Empty).Trim();
            var strB = (y ?? string.Empty).Trim();
            return string.CompareOrdinal(strA, strB);
        }
    }

    #endregion

    #region LoadDictionaries

    private protected static void LoadDictionaries()
    {
        if (OKSM.Count == 0)
        {
#if DEBUG
            OKSM_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\..\..\")), "data", "Spravochniki", "oksm.xlsx"));
#else
            OKSM_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"oksm.xlsx"));
#endif
        }
        if (R.Count == 0)
        {
#if DEBUG
            R_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\..\..\")), "data", "Spravochniki", "R.xlsx"));
#else
            R_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"R.xlsx"));
#endif
        }
        if (HolidaysSpecific.Count == 0)
        {
#if DEBUG
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\..\..\")), "data", "Spravochniki", "Holidays.xlsx"));
#else
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"Holidays.xlsx"));
#endif
        }
    }

    #region  HolidaysFromFile

    //функция импорта праздничных и рабочих дат из справочника
    private protected static void Holidays_Populate_From_File(string fileAddress)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        if (!File.Exists(fileAddress)) return;
        FileInfo excelImportFile = new(fileAddress);
        var xls = new ExcelPackage(excelImportFile);
        var worksheet1 = xls.Workbook.Worksheets["Выходные"];
        var i = 1;
        HolidaysSpecific.Clear();
        while (worksheet1.Cells[i, 1].Text != string.Empty)
        {
            var valueDate = worksheet1.Cells[i, 1].Text;
            if (DateOnly.TryParse(valueDate, out var fixDate))
            {
                HolidaysSpecific.Add(fixDate);
            }
            i++;
        }
        var worksheet2 = xls.Workbook.Worksheets["Рабочие"];
        i = 1;
        WorkDaysSpecific.Clear();
        while (worksheet2.Cells[i, 1].Text != string.Empty)
        {
            var valueDate = worksheet2.Cells[i, 1].Text;
            if (DateOnly.TryParse(valueDate, out var fixDate))
            {
                WorkDaysSpecific.Add(fixDate);
            }
            i++;
        }
    }

    #endregion

    #region OKSMFromFile

    private protected static void OKSM_Populate_From_File(string fileAddress)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        if (!File.Exists(fileAddress)) return;
        FileInfo excel_import_file = new(fileAddress);
        var xls = new ExcelPackage(excel_import_file);
        var worksheet1 = xls.Workbook.Worksheets["Лист1"];
        var i = 8;
        OKSM.Clear();
        while (worksheet1.Cells[i, 1].Text != string.Empty)
        {
            OKSM.Add(new Dictionary<string, string>
            {
                {"kod", worksheet1.Cells[i, 2].Text},
                {"shortname", worksheet1.Cells[i, 3].Text},
                {"longname", worksheet1.Cells[i, 4].Text},
                {"alpha2", worksheet1.Cells[i, 5].Text},
                {"alpha3", worksheet1.Cells[i, 6].Text}
            });
            i++;
        }
    }

    #endregion

    #region RFromFile

    private protected static void R_Populate_From_File(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        if (!File.Exists(filePath)) return;
        FileInfo excelImportFile = new(filePath);
        var xls = new ExcelPackage(excelImportFile);
        var worksheet = xls.Workbook.Worksheets["Лист1"];
        var i = 2;
        R.Clear();
        while (worksheet.Cells[i, 1].Text != string.Empty)
        {
            R.Add(new Dictionary<string, string>
            {
                {"name", worksheet.Cells[i, 1].Text},
                {"value", worksheet.Cells[i, 5].Text},
                {"unit", worksheet.Cells[i, 6].Text},
                {"code", worksheet.Cells[i, 8].Text},
                {"D", worksheet.Cells[i, 15].Text},
                {"MZA", worksheet.Cells[i, 17].Text},
                {"A_Solid", worksheet.Cells[i, 18].Text},
                {"A_Liquid", worksheet.Cells[i, 20].Text},
                {"OSPORB_Solid", worksheet.Cells[i, 22].Text},
                {"OSPORB_Liquid", worksheet.Cells[i, 23].Text}
            });
            if (string.IsNullOrWhiteSpace(R[^1]["D"]) || !double.TryParse(R[^1]["D"], out var val1) || val1 < 0)
            {
                R[^1]["D"] = double.MaxValue.ToString();
            }
            i++;
        }
    }

    #endregion

    #endregion

    #region OverdueCalculations
    //вычисление нарушения сроков предоставления отчётов с учетом праздников

    //кол-во дней между датой операции (документа для 10) и окончанием ОП

    protected static readonly Dictionary<string, int> OverduePeriods_RV = new()
    {

        { "10", 10 },{ "11", 10 },{ "12", 10 },                          { "15", 10 },             { "17", 10 },{ "18", 10 },
                     { "21", 10 },{ "22", 10 },                          { "25", 10 },             { "27", 10 },{ "28", 10 },{ "29", 10 },
                     { "31", 10 },{ "32", 10 },                          { "35", 10 },             { "37", 10 },{ "38", 10 },{ "39", 10 },
                     { "41", 10 },{ "42", 10 },{ "43", 10 },                          { "46", 10 },{ "47", 10 },{ "48", 10 },
                                               { "53", 10 },{ "54", 10 },                                       { "58", 10 },
                     { "61", 10 },{ "62", 10 },{ "63", 10 },{ "64", 10 },{ "65", 10 },{ "66", 10 },{ "67", 10 },{ "68", 10 },
                     { "71", 01 },{ "72", 10 },{ "73", 05 },{ "74", 05 },{ "75", 05 },
                     { "81", 10 },{ "82", 10 },{ "83", 10 },{ "84", 10 },{ "85", 10 },{ "86", 10 },{ "87", 10 },{ "88", 10 },
                                                                                                   { "97", 10 },{ "98", 10 },{ "99", 10 }
    };

    protected static readonly Dictionary<string, int> OverduePeriods_RAO = new()
    {
        { "01", 90 },
        { "10", 10 },{ "11", 10 },{ "12", 10 },{ "13", 10 },{ "14", 10 },             { "16", 10 },             { "18", 10 },
                     { "21", 10 },{ "22", 10 },                          { "25", 10 },{ "26", 10 },{ "27", 10 },{ "28", 10 },{ "29", 10 },
                     { "31", 10 },{ "32", 10 },                          { "35", 10 },{ "36", 10 },{ "37", 10 },{ "38", 10 },{ "39", 10 },
                     { "41", 10 },{ "42", 10 },{ "43", 10 },{ "44", 10 },{ "45", 10 },                          { "48", 10 },{ "49", 10 },
                     { "51", 10 },{ "52", 10 },                          { "55", 10 },{ "56", 10 },{ "57", 10 },             { "59", 10 },
                                               { "63", 10 },{ "64", 10 },                                       { "68", 10 },
                     { "71", 01 },{ "72", 10 },{ "73", 05 },{ "74", 05 },{ "75", 05 },{ "76", 10 },
                                                            { "84", 10 },                                       { "88", 10 },
                                                                                                   { "97", 10 },{ "98", 10 },{ "99", 10 }
    };

    //праздники из года в год
    private static readonly List<DateOnly> HolidaysGeneric = new()
    {
        new DateOnly(1,1,1), //1 января
        new DateOnly(1,1,2), //2 января
        new DateOnly(1,1,3), //3 января
        new DateOnly(1,1,4), //4 января
        new DateOnly(1,1,5), //5 января
        new DateOnly(1,1,6), //6 января
        new DateOnly(1,1,7), //7 января
        new DateOnly(1,1,8), //8 января
        new DateOnly(1,2,23), //23 февраля
        new DateOnly(1,3,8), //8 марта
        new DateOnly(1,5,1), //1 мая
        new DateOnly(1,5,9), //9 мая
        new DateOnly(1,6,12), //12 июня
        new DateOnly(1,11,4), //4 ноября
    };

    //рабочие дни по выходным, уникальные для каждого года, берутся из файла ./Spravochniki/Holidays.xlsx
    private static readonly List<DateOnly> WorkDaysSpecific = new();

    //праздники конкретных годов, берутся из файла ./Spravochniki/Holidays.xlsx
    private protected static readonly List<DateOnly> HolidaysSpecific = new();

    //расчет кол-ва рабочих дней между двумя датами
    //strict_order - ожидать даты в правильном порядке
    private protected static int WorkdaysBetweenDates(DateOnly date1, DateOnly date2, bool strictOrder = true)
    {
        var result = 0;
        DateOnly dateMin;
        DateOnly dateMax;
        if (strictOrder)
        {
            dateMin = date1;
            dateMax = date2;
        }
        else
        {
            dateMin = date1 < date2 ? date1 : date2;
            dateMax = dateMin == date2 ? date1 : date2;
        }
        if (dateMin > dateMax)
        {
            return int.MaxValue;
        }
        for (var day = dateMin.AddDays(1); day <= dateMax; day = day.AddDays(1))
        {
            var isWorkingDay = WorkDaysSpecific.Any(x => Equals(x, day))
                               || !(day.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday
                                    || HolidaysSpecific.Any(x => Equals(x, day))
                                    || HolidaysGeneric.Any(x => Equals(x.Month, day.Month) && Equals(x.Day, day.Day)));
            if (isWorkingDay)
            {
                result++;
            }
        }
        return result;
    }

    #endregion

    #region TryParseFloatExtended

    protected static bool TryParseFloatExtended(string? str, out float val)
    {
        return float.TryParse(ConvertStringToExponential(str),
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
            CultureInfo.CreateSpecificCulture("ru-RU"),
            out val);
    }

    #endregion
}
