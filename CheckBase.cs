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
    private protected static Dictionary<string, double> D = new();

    private protected static List<Dictionary<string, string>> OKSM = new();

    private protected static List<Dictionary<string, string>> R = new();

    private protected static bool DB_Ignore = true;

    #region CheckNotePresence

    private protected static bool CheckNotePresence(List<Form> forms, List<Note> notes, int line, byte graphNumber)
    {
        var valid = false;
        foreach (var note in notes)
        {
            if (note.RowNumber_DB == null || forms[line].ReportId == null) continue;
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

    #region DFromFile

    private protected static void D_Populate_From_File(string file_address)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        if (!File.Exists(file_address)) return;
        FileInfo excel_import_file = new(file_address);
        var xls = new ExcelPackage(excel_import_file);
        var worksheet1 = xls.Workbook.Worksheets["Лист1"];
        var i = 2;
        string name_1, name_2, name_base, name_real;
        name_base = "аврорий";
        D.Clear();
        while (worksheet1.Cells[i, 1].Text != string.Empty)
        {
            name_1 = worksheet1.Cells[i, 2].Text;
            name_2 = worksheet1.Cells[i, 3].Text;
            if (name_1 != string.Empty)
            {
                name_base = name_1.ToLower();
            }
            if (name_2.Contains('-'))
            {
                if (name_2.Contains('+'))
                {
                    name_real = name_base + name_2[name_2.IndexOf('-')..name_2.IndexOf('+')];
                }
                else
                {
                    name_real = name_base + name_2[name_2.IndexOf('-')..];
                }
                var valueBase = worksheet1.Cells[i, 4].Text;
                double valueReal;
                if (valueBase.Contains("Неограниченно"))
                {
                    valueReal = double.MaxValue;
                }
                else
                {
                    valueReal = 1e12 * double.Parse(valueBase[..6].Replace(" ", ""), NumberStyles.Float);
                }
                D[name_real] = valueReal;
                if (name_real.Contains("йод"))
                {
                    D[name_real.Replace('й', 'и')] = valueReal;
                }
                else if (name_real.Contains("иод"))
                {
                    D[name_real.Replace('и', 'й')] = valueReal;
                }
            }
            i++;
        }
    }

    #endregion

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
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
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
                {"D", worksheet.Cells[i, 15].Text}

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
}
