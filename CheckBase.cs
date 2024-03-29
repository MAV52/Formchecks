using System.Collections.Generic;
using System.Linq;
using Models.Forms;
using OfficeOpenXml;
using System.Globalization;
using System.IO;

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

    #region DFromFile

    private protected static void D_Populate_From_File(string file_address)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        if (!File.Exists(file_address)) return;
        FileInfo excel_import_file = new(file_address);
        var xls = new ExcelPackage(excel_import_file);
        var wrksht1 = xls.Workbook.Worksheets["Лист1"];
        var i = 2;
        string name_1, name_2, name_base, name_real;
        name_base = "аврорий";
        string value_base;
        double value_real;
        D.Clear();
        while (wrksht1.Cells[i, 1].Text != string.Empty)
        {
            name_1 = wrksht1.Cells[i, 2].Text;
            name_2 = wrksht1.Cells[i, 3].Text;
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
                value_base = wrksht1.Cells[i, 4].Text;
                if (value_base.Contains("Неограниченно"))
                {
                    value_real = double.MaxValue;
                }
                else
                {
                    value_real = 1e12 * double.Parse(value_base[..6].Replace(" ", ""), NumberStyles.Float);
                }
                D[name_real] = value_real;
                if (name_real.Contains("йод"))
                {
                    D[name_real.Replace('й', 'и')] = value_real;
                }
                else if (name_real.Contains("иод"))
                {
                    D[name_real.Replace('и', 'й')] = value_real;
                }
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
        var wrksht1 = xls.Workbook.Worksheets["Лист1"];
        var i = 8;
        OKSM.Clear();
        while (wrksht1.Cells[i, 1].Text != string.Empty)
        {
            OKSM.Add(new Dictionary<string, string>
            {
                {"kod", wrksht1.Cells[i, 2].Text},
                {"shortname", wrksht1.Cells[i, 3].Text},
                {"longname", wrksht1.Cells[i, 4].Text},
                {"alpha2", wrksht1.Cells[i, 5].Text},
                {"alpha3", wrksht1.Cells[i, 6].Text}
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
                {"unit", worksheet.Cells[i, 6].Text}
            });
            i++;
        }
    }

    #endregion
}