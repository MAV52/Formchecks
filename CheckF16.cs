using Models.CheckForm;
using Models.Collections;
using Models.Forms;
using Models.Forms.Form1;
using OfficeOpenXml.Drawing.Vml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Client_App.Commands.SyncCommands.CheckForm;

public class CheckF16 : CheckBase
{
    #region Properties

    private static readonly string[] OperationCode_DB_Valids =
    {
        "10","11","12","13","14","16","18",
        "21","22","25","26","27","28","29",
        "31","32","35","36","37","38","39",
        "42","43","44","45","48","49","51",
        "52","56","57","59","63","64","68",
        "71","72","73","74","75","76","97",
        "98","99"
    };

    private static readonly Dictionary<string, string> GraphsList = new()
    {
        { "NumberInOrder_DB", "01 - № п/п" },
        { "OperationCode_DB", "02 - Код операции" },
        { "OperationDate_DB", "03 - Дата операции" },
        { "CodeRAO_DB", "04 - Код РАО" },
        { "StatusRAO_DB", "05 - Статус РАО" },
        { "Volume_DB", "06 - Объем без упаковки, куб. м" },
        { "Mass_DB", "07 - масса без упаковки (нетто), т" },
        { "QuantityOZIII_DB", "08 - Количество ОЗИИИ, шт." },
        { "MainRadionuclids_DB", "09 - Основные радионуклиды" },
        { "TritiumActivity_DB", "10 - Суммарная активность, Бк: тритий" },
        { "BetaGammaActivity_DB", "11 - Суммарная активность, Бк: бета-, гамма-излучающие радионуклиды" },
        { "AlphaActivity_DB", "12 - Суммарная активность, Бк: альфа-излучающие радионуклиды" },
        { "TransuraniumActivity_DB", "13 - Суммарная активность, Бк: трансурановые радионуклиды" },
        { "ActivityMeasurementDate_DB", "14 - Дата измерения активности" },
        { "DocumentVid_DB", "15 - Вид документа" },
        { "DocumentNumber_DB", "16 - Номер документа" },
        { "DocumentDate_DB", "17 - Дата документа" },
        { "ProviderOrRecieverOKPO_DB", "18 - Код ОКПО поставщика или получателя" },
        { "TransporterOKPO_DB", "19 - Код ОКПО перевозчика" },
        { "StoragePlaceName_DB", "20 - Наименование пункта хранения" },
        { "StoragePlaceCode_DB", "21 - Код пункта хранения" },
        { "RefineOrSortRAOCode_DB", "22 - Код переработки/сортировки РАО" },
        { "PackName_DB", "23 - Наименование УКТ, упаковки или иной ученой единицы" },
        { "PackType_DB", "24 - Тип УКТ, упаковки или иной ученой единицы" },
        { "PackNumber_DB", "25 - Заводской номер УКТ, упаковки или иной ученой единицы" },
        { "Subsidy_DB", "26 - Субсидия, %" },
        { "FcpNumber_DB", "27 - Номер мероприятия ФЦП" },
    };

    #endregion

    #region CheckTotal

    public static List<CheckError> Check_Total(Reports reps, Report rep)
    {
        var currentFormLine = 0;
        List<CheckError> errorList = new();
        if (OKSM.Count == 0)
        {
#if DEBUG
            OKSM_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\")), "data", "Spravochniki", "oksm.xlsx"));
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
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\")), "data", "Spravochniki", "Holidays.xlsx"));
#else
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"Holidays.xlsx"));
#endif
        }
        var formsList = rep.Rows16.ToList<Form16>();
        foreach (var key in rep.Rows16)
        {
            var form = (Form16)key;
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_001(formsList, currentFormLine));
            errorList.AddRange(Check_002(formsList, currentFormLine));
            errorList.AddRange(Check_002_11(formsList, currentFormLine));
            errorList.AddRange(Check_002_12(formsList, currentFormLine));
            errorList.AddRange(Check_002_29(formsList, notes, currentFormLine));
            errorList.AddRange(Check_002_51(formsList, currentFormLine));
            errorList.AddRange(Check_002_52(formsList, currentFormLine));
            errorList.AddRange(Check_002_57(formsList, currentFormLine));
            errorList.AddRange(Check_002_10(formsList, currentFormLine));
            errorList.AddRange(Check_002_49(formsList, currentFormLine));
            errorList.AddRange(Check_002_71(formsList, currentFormLine));
            errorList.AddRange(Check_003_non10(formsList, rep, currentFormLine));
            errorList.AddRange(Check_003_10(formsList, rep, currentFormLine));
            errorList.AddRange(Check_004(formsList, notes, currentFormLine));
            errorList.AddRange(Check_005_11(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_005_76(formsList, currentFormLine));
            errorList.AddRange(Check_005_Other(formsList, currentFormLine));
            errorList.AddRange(Check_006(formsList, currentFormLine));
            errorList.AddRange(Check_007(formsList, currentFormLine));
            errorList.AddRange(Check_008(formsList, currentFormLine));
            errorList.AddRange(Check_009(formsList, currentFormLine));
            errorList.AddRange(Check_010(formsList, currentFormLine));
            errorList.AddRange(Check_011(formsList, currentFormLine));
            errorList.AddRange(Check_012(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_013(formsList, currentFormLine));
            errorList.AddRange(Check_014(formsList, currentFormLine));
            errorList.AddRange(Check_015(formsList, currentFormLine));
            errorList.AddRange(Check_016(formsList, notes, currentFormLine));
            errorList.AddRange(Check_017(formsList, rep, currentFormLine));
            errorList.AddRange(Check_018_10(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_018_21(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_018_22(formsList, currentFormLine));
            errorList.AddRange(Check_019_01(formsList, currentFormLine));
            errorList.AddRange(Check_019_21(formsList, currentFormLine));
            errorList.AddRange(Check_019_22(formsList, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, currentFormLine));
            errorList.AddRange(Check_022_44(formsList, currentFormLine));
            errorList.AddRange(Check_022_45(formsList, currentFormLine));
            errorList.AddRange(Check_022_49(formsList, currentFormLine));
            errorList.AddRange(Check_022_10(formsList, currentFormLine));
            errorList.AddRange(Check_022_22_44(formsList, currentFormLine));
            errorList.AddRange(Check_022_22_56(formsList, currentFormLine));
            errorList.AddRange(Check_023(formsList, currentFormLine));
            errorList.AddRange(Check_024(formsList, currentFormLine));
            errorList.AddRange(Check_025(formsList, currentFormLine));
            errorList.AddRange(Check_026(formsList, currentFormLine));
            errorList.AddRange(Check_027(formsList, currentFormLine));
            currentFormLine++;
        }
        errorList.AddRange(Check_028(formsList));
        errorList.AddRange(Check_029(formsList, rep));
        var index = 0;
        foreach (var error in errorList)
        {
            if (GraphsList.TryGetValue(error.Column, out var columnFrontName))
            {
                error.Column = columnFrontName;
            }
            index++;
            error.Index = index;
        }
        return errorList;
    }

    #endregion

    #region Checks

    #region Check001

    private static List<CheckError> Check_001(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var valid = (line == 0 && forms[line].NumberInOrder_DB is 1 or 0) || line > 0 && forms[line - 1].NumberInOrder_DB == forms[line].NumberInOrder_DB - 1;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "NumberInOrder_DB",
                Value = forms[line].NumberInOrder_DB.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Номера строк должны располагаться по порядку, без пропусков или дублирования номеров"
            });
        }
        return result;
    }

    #endregion

    #region Check002

    private static List<CheckError> Check_002(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var valid = operationCode != null && OperationCode_DB_Valids.Contains(operationCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Код операции не может быть использован в форме 1.6."
            });
        }
        return result;
    }

    #endregion

    #region Check002_11

    private static List<CheckError> Check_002_11(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "11", "12", "13", "14", "16", "18", "41", "56", "57", "59" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        if (forms[line].CodeRAO_DB.Length < 10) return result;
        var valid = forms[line].CodeRAO_DB.Substring(8, 2) != "99";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для вновь образованных РАО код типа РАО «99» не может быть использован"
            });
        }
        return result;
    }

    #endregion

    #region Check002_12

    private static List<CheckError> Check_002_12(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "12,42" };
        var requiredNuclids = new string[] { "плутоний", "уран-233", "уран-235", "уран-238", "нептуний-237", "америций-241", "америций-243", "калифорний-252", "торий", "литий-6", "тритий" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var nuclids = forms[line].MainRadionuclids_DB.Split(';');
        var valid = false;
        for (var i = 0; i < nuclids.Length; i++)
        {
            nuclids[i] = nuclids[i].Trim().ToLower();
            if (requiredNuclids.Contains(nuclids[i]))
            {
                valid = true;
                break;
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "MainRadionuclids_DB",
                Value = forms[line].MainRadionuclids_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В графе 9 не представлены сведения о радионуклидах, которые могут быть отнесены к ЯМ. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_29

    private static List<CheckError> Check_002_29(List<Form16> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "29", "39", "49", "59", "97", "98", "99" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        const byte graphNumber = 2;
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо дать пояснение об осуществленной операции."
            });
        }
        return result;
    }

    #endregion

    #region Check002_51

    private static List<CheckError> Check_002_51(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "51" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + ""
            });
        }
        return result;
    }

    #endregion

    #region Check002_52

    private static List<CheckError> Check_002_52(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "52" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В отчетах не найдена строка об осуществлении передачи учетной единицы. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_57
    private static List<CheckError> Check_002_57(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "57" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В отчетах не найдена строка снятии учетной единицы для упаковки/переупаковки. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_10
    private static List<CheckError> Check_002_10(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "10", "11", "12", "13", "14", "16", "18", "31", "32", "35", "36", "37", "38", "41", "56", "57", "59", "97" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var mass = forms[line].Mass_DB;
        var mass_exists = TryParseFloatExtended(mass, out var mass_real);
        var activity_a = forms[line].AlphaActivity_DB;
        var activity_b = forms[line].BetaGammaActivity_DB;
        var activity_a_exists = TryParseFloatExtended(activity_a, out var activity_a_real);
        var activity_b_exists = TryParseFloatExtended(activity_b, out var activity_b_real);
        var CodeRAO_1_MatterState = forms[line].CodeRAO_DB.Substring(0, 1);
        var valid = true;
        if (mass_exists && CodeRAO_1_MatterState is "1" or "2" && (activity_a_exists || activity_b_exists))
        {
            if (CodeRAO_1_MatterState == "1")
            {
                valid = (activity_a_exists && activity_a_real / (mass_real * 1e6) >= 0.05f) || (activity_b_exists && activity_b_real / (mass_real * 1e6) >= 0.5f);
            }
            else if (CodeRAO_1_MatterState == "2")
            {
                valid = (activity_a_exists && activity_a_real / (mass_real * 1e6) >= 1.0f) || (activity_b_exists && activity_b_real / (mass_real * 1e6) >= 100.0f);
            }
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "По сведениям, представленным в строке, отходы не относятся к РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check002_49
    private static List<CheckError> Check_002_49(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "49", "59" };
        var graph22ValidValues = new string[] { "-", "52", "72", "74" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = graph22ValidValues.Contains(forms[line].RefineOrSortRAOCode_DB.Trim());
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Код операции не соответствует коду переработки/сортировки, указанному в графе 22"
            });
        }
        return result;
    }

    #endregion

    #region Check002_71
    //Справочная "ошибка" - т.е. даже не ошибка.
    private static List<CheckError> Check_002_71(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "71", "72", "73", "74", "75", "76", "18", "68" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "К отчету необходимо приложить скан-копию документа характеризующего операцию"
            });
        }
        return result;
    }

    #endregion

    #region Check003_non10
    private static List<CheckError> Check_003_non10(List<Form16> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] nonApplicableOperationCodes = { "10" };
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        if (nonApplicableOperationCodes.Contains(operationCode)) return result;
        var valid = operationDate != null;
        DateTime pEnd;
        DateTime pMid;
        if (valid && rep is { StartPeriod_DB: not null, EndPeriod_DB: not null })
        {
            valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(operationDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(operationDate),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата операции не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion

    #region Check003_10
    private static List<CheckError> Check_003_10(List<Form16> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        var operationCode = forms[line].OperationCode_DB;
        var documentDate = forms[line].DocumentDate_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = documentDate != null;
        DateTime pEnd;
        DateTime pMid;
        if (valid && rep is { StartPeriod_DB: not null, EndPeriod_DB: not null })
        {
            valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(documentDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(documentDate),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа не входит в отчетный период. Для операции инвентаризации срок предоставления отчета исчисляется с даты утверждения акта инвентаризации"
            });
        }
        return result;
    }

    #endregion

    #region Check004
    private static List<CheckError> Check_004(List<Form16> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var comparator = new CustomNullStringWithTrimComparer();
        var CodeRAO_DB = forms[line].CodeRAO_DB.Trim();
        var valid = !string.IsNullOrWhiteSpace(CodeRAO_DB) && CodeRAO_DB.Length == 11 && CodeRAO_DB.All(char.IsDigit);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность заполнения кода РАО."
            });
            return result;
        }

        #region data fetch

        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        var halflife_max = 0.0f;
        var halflife_max_id = -1;
        var validUnits = new Dictionary<string, float>()
            {
                { "лет", 1f },
                { "сут", 365.242374f },
                { "час", 365.242374f*24.0f },
                { "мин", 365.242374f*24.0f*60.0f },
                { "сек", 365.242374f*24.0f*60.0f*60.0f }
            };
        var nuclid_activity_t = forms[line].TritiumActivity_DB;
        var nuclid_activity_a = forms[line].AlphaActivity_DB;
        var nuclid_activity_b = forms[line].BetaGammaActivity_DB;
        var nuclid_activity_u = forms[line].TransuraniumActivity_DB;
        var nuclid_mass = forms[line].Mass_DB;
        var nuclids_exist_t = TryParseFloatExtended(nuclid_activity_t, out float nuclid_activity_t_real);
        var nuclids_exist_a = TryParseFloatExtended(nuclid_activity_a, out float nuclid_activity_a_real);
        var nuclids_exist_b = TryParseFloatExtended(nuclid_activity_b, out float nuclid_activity_b_real);
        var nuclids_exist_u = TryParseFloatExtended(nuclid_activity_u, out float nuclid_activity_u_real);
        var nuclid_mass_exists = TryParseFloatExtended(nuclid_mass, out float nuclid_mass_real);
        foreach (string nuclid in radArray)
        {
            var nuclid_id = R.FindIndex(x => comparator.Compare(x["name"], nuclid) == 0);
            if (nuclid_id >= 0 && TryParseFloatExtended(R[nuclid_id]["value"], out var halflife_val) && validUnits.ContainsKey(R[nuclid_id]["unit"]))
            {
                halflife_max = Math.Max(halflife_val / validUnits[R[nuclid_id]["unit"]], halflife_max);
                halflife_max_id = nuclid_id;
            }
        }
        const byte graphNumber = 4;
        var noteExists = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);

        #endregion

        #region setup

        var CodeRAO_1_MatterState = CodeRAO_DB.Substring(0, 1);
        var CodeRAO_2_RAOCategory = CodeRAO_DB.Substring(1, 1);
        var CodeRAO_3_NuclidTypes = CodeRAO_DB.Substring(2, 1);
        var CodeRAO_4_HasNuclears = CodeRAO_DB.Substring(3, 1);
        var CodeRAO_5_HalfLife = CodeRAO_DB.Substring(4, 1);
        var CodeRAO_6_DangerPeriod = CodeRAO_DB.Substring(5, 1);
        var CodeRAO_7_RecycleMethod = CodeRAO_DB.Substring(6, 1);
        var CodeRAO_8_RAOClass = CodeRAO_DB.Substring(7, 1);
        var CodeRAO_910_TypeCode = CodeRAO_DB.Substring(8, 2);
        var CodeRAO_11_Flammability = CodeRAO_DB.Substring(10, 1);
        var CodeRAO_1_Allowed = new string[] { "1", "2", "3" };
        var CodeRAO_2_Allowed = new string[] { "0", "1", "2", "3", "4", "9" };
        var CodeRAO_3_Allowed = new string[] { "1", "2", "3", "4", "5", "6" };
        var CodeRAO_4_Allowed = new string[] { "1", "2" };
        var CodeRAO_5_Allowed = new string[] { "1", "2" };
        var CodeRAO_6_Allowed = new string[] { "0", "1", "2", "3" };
        var CodeRAO_7_Allowed = new string[] { "0", "1", "2", "3", "4", "9" };
        var CodeRAO_8_Allowed = new string[] { "0", "1", "2", "3", "4", "5", "6", "7", "9" };
        var CodeRAO_910_Allowed = new string[]
        {
            "01",
            "11","12","13","14","15","16","17","18","19",
            "21","22","23","24","25","26",          "29",
            "31","32","33","34","35","36","37","38","39",
            "41","42","43","44","45","46",
            "51","52","53","54","55","56","57","58","59",
            "61","62","63","64","65","66","67","68","69",
            "71","72","73","74","75","76","77","78","79",
            "81","82","83","84","85","86","87","88","89",
            "91","92","93","94","95","96","97","98","99"
        };
        var CodeRAO_11_Allowed = new string[] { "1", "2" };

        var CodeRAO_1_Valid = CodeRAO_1_Allowed.Contains(CodeRAO_1_MatterState);
        var CodeRAO_2_Valid = CodeRAO_2_Allowed.Contains(CodeRAO_2_RAOCategory);
        var CodeRAO_3_Valid = CodeRAO_3_Allowed.Contains(CodeRAO_3_NuclidTypes);
        var CodeRAO_4_Valid = CodeRAO_4_Allowed.Contains(CodeRAO_4_HasNuclears);
        var CodeRAO_5_Valid = CodeRAO_5_Allowed.Contains(CodeRAO_5_HalfLife);
        var CodeRAO_6_Valid = CodeRAO_6_Allowed.Contains(CodeRAO_6_DangerPeriod);
        var CodeRAO_7_Valid = CodeRAO_7_Allowed.Contains(CodeRAO_7_RecycleMethod);
        var CodeRAO_8_Valid = CodeRAO_8_Allowed.Contains(CodeRAO_8_RAOClass);
        var CodeRAO_910_Valid = CodeRAO_910_Allowed.Contains(CodeRAO_910_TypeCode);
        var CodeRAO_11_Valid = CodeRAO_11_Allowed.Contains(CodeRAO_11_Flammability);

        #endregion

        #region symbol 1

        if (!CodeRAO_1_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_1_MatterState,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 1-го символа кода РАО."
            });
        }
        else
        {
            var validTypeCode_Liquid = new List<string>();
            var validTypeCode_Solid = new List<string>();
            var validTypeCode_Solid_7Not0 = new List<string>();
            var validTypeCode_Gaseous = new List<string>();
            //liquids: symbols 9-10 are 11-39, 99.
            for (int i = 11; i <= 39; i++)
            {
                validTypeCode_Liquid.Add(i.ToString("D2"));
            }
            validTypeCode_Liquid.Add(99.ToString("D2"));
            //solids: symbols 9-10 are 41-99; 11-39 if symbol 7 is not 0.
            for (int i = 41; i <= 99; i++)
            {
                validTypeCode_Solid.Add(i.ToString("D2"));
            }
            for (int i = 11; i <= 39; i++)
            {
                validTypeCode_Solid_7Not0.Add(i.ToString("D2"));
            }
            //gases: symbols 9-10 are strictly 01.
            for (int i = 1; i <= 1; i++)
            {
                validTypeCode_Gaseous.Add(i.ToString("D2"));
            }
            switch (CodeRAO_1_MatterState)
            {
                case "1":
                    if (!validTypeCode_Liquid.Contains(CodeRAO_910_TypeCode))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_16",
                            Row = forms[line].NumberInOrder_DB.ToString(),
                            Column = "CodeRAO_DB",
                            Value = CodeRAO_DB,
                            Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Агрегатное состояние (символ №1) не соответствует типу выбранных РАО (символы №9-10)."
                        });
                    }
                    break;
                case "2":
                    if (!(validTypeCode_Solid.Contains(CodeRAO_910_TypeCode) || (validTypeCode_Solid_7Not0.Contains(CodeRAO_910_TypeCode) && CodeRAO_7_RecycleMethod != "0")))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_16",
                            Row = forms[line].NumberInOrder_DB.ToString(),
                            Column = "CodeRAO_DB",
                            Value = CodeRAO_DB,
                            Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Агрегатное состояние (символ №1) не соответствует типу выбранных РАО (символы №9-10)."
                        });
                    }
                    break;
                case "3":
                    if (!validTypeCode_Gaseous.Contains(CodeRAO_910_TypeCode))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_16",
                            Row = forms[line].NumberInOrder_DB.ToString(),
                            Column = "CodeRAO_DB",
                            Value = CodeRAO_DB,
                            Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Агрегатное состояние (символ №1) не соответствует типу выбранных РАО (символы №9-10)."
                        });
                    }
                    break;
            }
        }
        #endregion

        #region symbol 2

        if (!CodeRAO_2_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_2_RAOCategory,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 2-го символа кода РАО."
            });
        }
        else
        {
            switch (CodeRAO_2_RAOCategory)
            {
                case "4":
                    var validTypeCode = new string[] { "81", "82", "85", "86", "87", "88", "89" };
                    if (!validTypeCode.Contains(CodeRAO_910_TypeCode))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_16",
                            Row = forms[line].NumberInOrder_DB.ToString(),
                            Column = "CodeRAO_DB",
                            Value = CodeRAO_2_RAOCategory,
                            Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Значение 2-го символа кода РАО 4 используется только для отработавших ЗРИ."
                        });
                    }
                    break;
                case "9":
                    if (!noteExists)
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_16",
                            Row = forms[line].NumberInOrder_DB.ToString(),
                            Column = "CodeRAO_DB",
                            Value = CodeRAO_2_RAOCategory,
                            Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо дать пояснение для 2-го символа кода РАО."
                        });
                    }
                    break;
                default:
                    // 0, 1, 2, 3, 9
                    if (CodeRAO_2_RAOCategory == "0" && CodeRAO_1_MatterState == "1")
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_16",
                            Row = forms[line].NumberInOrder_DB.ToString(),
                            Column = "CodeRAO_DB",
                            Value = CodeRAO_2_RAOCategory,
                            Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Неправильно указана категория РАО."
                        });
                    }
                    else
                    {
                        var code_max = -1;
                        if (nuclid_mass_exists && nuclid_mass_real > 0)
                        {
                            #region A for tritium
                            if (nuclids_exist_t)
                            {
                                var A = nuclid_activity_t_real / (nuclid_mass_real * 1e6);
                                if (CodeRAO_1_MatterState == "2")
                                {
                                    if (A < 1e07) { if (code_max < 0) code_max = 0; }
                                    else if (A < 1e08) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e11) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                                else if (CodeRAO_1_MatterState == "1")
                                {
                                    if (A < 1e04) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e08) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                            }
                            #endregion
                            #region A for beta-gamma
                            if (nuclids_exist_b)
                            {
                                var A = nuclid_activity_b_real / (nuclid_mass_real * 1e6);
                                if (CodeRAO_1_MatterState == "2")
                                {
                                    if (A < 1e03) { if (code_max < 0) code_max = 0; }
                                    else if (A < 1e04) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e07) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                                else if (CodeRAO_1_MatterState == "1")
                                {
                                    if (A < 1e03) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e07) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                            }
                            #endregion
                            #region A for alpha
                            if (nuclids_exist_a)
                            {
                                var A = nuclid_activity_a_real / (nuclid_mass_real * 1e6);
                                if (CodeRAO_1_MatterState == "2")
                                {
                                    if (A < 1e02) { if (code_max < 0) code_max = 0; }
                                    else if (A < 1e03) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e06) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                                else if (CodeRAO_1_MatterState == "1")
                                {
                                    if (A < 1e02) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e06) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                            }
                            #endregion
                            #region A for transuraniums
                            if (nuclids_exist_u)
                            {
                                var A = nuclid_activity_u_real / (nuclid_mass_real * 1e6);
                                if (CodeRAO_1_MatterState == "2")
                                {
                                    if (A < 1e01) { if (code_max < 0) code_max = 0; }
                                    else if (A < 1e02) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e05) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                                else if (CodeRAO_1_MatterState == "1")
                                {
                                    if (A < 1e01) { if (code_max < 1) code_max = 1; }
                                    else if (A < 1e05) { if (code_max < 2) code_max = 2; }
                                    else { if (code_max < 3) code_max = 3; }
                                }
                            }
                            #endregion
                        }
                        if (code_max == -1 && CodeRAO_2_RAOCategory != "9")
                        {
                            result.Add(new CheckError
                            {
                                FormNum = "form_16",
                                Row = forms[line].NumberInOrder_DB.ToString(),
                                Column = "CodeRAO_DB",
                                Value = CodeRAO_2_RAOCategory,
                                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте категорию РАО и суммарную активность."
                            });
                        }
                        else if (code_max != -1 && (CodeRAO_2_RAOCategory == "9" || CodeRAO_2_RAOCategory != code_max.ToString("D1")))
                        {
                            result.Add(new CheckError
                            {
                                FormNum = "form_16",
                                Row = forms[line].NumberInOrder_DB.ToString(),
                                Column = "CodeRAO_DB",
                                Value = CodeRAO_2_RAOCategory,
                                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + $"По данным, представленным в строке {forms[line].NumberInOrder_DB}, категория РАО {code_max}."
                            });
                        }
                    }
                    break;
            }
        }
        #endregion

        #region symbol 3

        if (!CodeRAO_3_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_3_NuclidTypes,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 3-го символа кода РАО."
            });
        }
        else
        {
            if (radArray.Length == 0)
            {
                if (CodeRAO_3_NuclidTypes != "0")
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_16",
                        Row = forms[line].NumberInOrder_DB.ToString(),
                        Column = "CodeRAO_DB",
                        Value = CodeRAO_3_NuclidTypes,
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При отсутствии радионуклидов в графе 9 3-й символ кода РАО должен быть равен 0."
                    });
                }
            }
            else
            {
                var contains_b = nuclids_exist_b && radArray.Any(x => R.Any(y => comparator.Compare(y["name"], x) == 0 && comparator.Compare(y["code"], "б") == 0));
                var contains_a = nuclids_exist_a && radArray.Any(x => R.Any(y => comparator.Compare(y["name"], x) == 0 && comparator.Compare(y["code"], "а") == 0));
                var contains_u = nuclids_exist_u && radArray.Any(x => R.Any(y => comparator.Compare(y["name"], x) == 0 && comparator.Compare(y["code"], "у") == 0));
                var contains_t = nuclids_exist_t && radArray.Any(x => R.Any(y => comparator.Compare(y["name"], x) == 0 && comparator.Compare(y["code"], "т") == 0));
                var expectedValue = "0";
                if (!contains_t && !contains_b && !contains_a && contains_u) expectedValue = "1";
                else if (!contains_t && !contains_b && contains_a && !contains_u) expectedValue = "2";
                else if (!contains_t && !contains_b && contains_a && contains_u) expectedValue = "3";
                else if (!contains_t && contains_b && !contains_a && !contains_u) expectedValue = "4";
                else if (contains_t && !contains_b && !contains_a && !contains_u) expectedValue = "4";
                else if (contains_t && contains_b && !contains_a && !contains_u) expectedValue = "4";
                else if (!contains_t && contains_b && contains_a && !contains_u) expectedValue = "5";
                else if (contains_t && contains_b && contains_a && !contains_u) expectedValue = "5";
                else if (!contains_t && contains_b && contains_a && contains_u) expectedValue = "6";
                else if (contains_t && contains_b && contains_a && contains_u) expectedValue = "6";
                if (expectedValue != CodeRAO_3_NuclidTypes)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_16",
                        Row = forms[line].NumberInOrder_DB.ToString(),
                        Column = "CodeRAO_DB",
                        Value = CodeRAO_3_NuclidTypes,
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Третий символ кода РАО не соответствует сведениям о суммарной активности (графы 10-13) и/или радионуклидам, указанным в графе 9"
                    });
                }
            }

        }
        #endregion

        #region symbol 4

        if (!CodeRAO_4_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_4_HasNuclears,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 4-го символа кода РАО."
            });
        }
        else
        {
            var nuclears = new string[] { "плутоний", "уран-233", "уран-235", "уран-238", "нептуний-237", "америций-241", "америций-243", "калифорний-252", "торий", "литий-6", "тритий" };
            if (CodeRAO_4_HasNuclears == "2" && !radArray.Any(x => nuclears.Contains(x.ToLower())))
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "CodeRAO_DB",
                    Value = CodeRAO_4_HasNuclears,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "4-ый символ кода РАО равен 2 только при указании радионуклидов, которые могут быть отнесены к ЯМ."
                });
            }
            else if (CodeRAO_4_HasNuclears == "1" && radArray.Any(x => nuclears.Contains(x.ToLower())))
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "CodeRAO_DB",
                    Value = CodeRAO_4_HasNuclears,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "4-ый символ кода РАО равен 1 только при отсутствии радионуклидов, которые могут быть отнесены к ЯМ."
                });
            }
        }
        #endregion

        #region symbol 5

        if (!CodeRAO_5_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_5_HalfLife,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 5-го символа кода РАО."
            });
        }
        else
        {
            if (CodeRAO_5_HalfLife != "2" && (int)halflife_max <= 31)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "CodeRAO_DB",
                    Value = CodeRAO_5_HalfLife,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + $"По данным, представленным в строке {forms[line].NumberInOrder_DB}, 5-ый символ кода РАО (период полураспада) должен быть равен 2."
                });
            }
            else if (CodeRAO_5_HalfLife != "1" && (int)halflife_max > 31)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "CodeRAO_DB",
                    Value = CodeRAO_5_HalfLife,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + $"По данным, представленным в строке {forms[line].NumberInOrder_DB}, 5-ый символ кода РАО (период полураспада) должен быть равен 1."
                });
            }
        }
        #endregion

        #region symbol 6

        if (!CodeRAO_6_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_6_DangerPeriod,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 6-го символа кода РАО."
            });
        }
        else if (CodeRAO_1_MatterState != "3")
        {
            if (CodeRAO_1_MatterState is "1" or "2")
            {
                if (CodeRAO_3_NuclidTypes == "0")
                {
                    if (!(CodeRAO_6_DangerPeriod == "0" && noteExists))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_16",
                            Row = forms[line].NumberInOrder_DB.ToString(),
                            Column = "CodeRAO_DB",
                            Value = CodeRAO_6_DangerPeriod,
                            Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Укажите причины невозможности определения периода потенциальной опасности."
                        });
                    }
                }
            }
            if (CodeRAO_1_MatterState is "1" or "2")
            {
                if (radArray.Length > 0)
                {
                    var soft_warning = radArray.Length > 1;
                    float A;
                    float expectedPeriod = -1.0f;
                    int expectedValue = 0;
                    foreach (string nuclid in radArray)
                    {
                        var nuclid_data = R.FirstOrDefault(x => comparator.Compare(x["name"], nuclid) == 0, []);
                        if (nuclid_data.Count > 0 && TryParseFloatExtended(nuclid_data["value"], out var T) && validUnits.TryGetValue(nuclid_data["unit"], out var unit_adjustment))
                        {
                            var nuclid_activity = -1.0f;
                            if (nuclid_data["code"] == "а" && nuclids_exist_a) { nuclid_activity = nuclid_activity_a_real; }
                            else if (nuclid_data["code"] == "б" && nuclids_exist_b) { nuclid_activity = nuclid_activity_b_real; }
                            else if (nuclid_data["code"] == "у" && nuclids_exist_u) { nuclid_activity = nuclid_activity_u_real; }
                            else if (nuclid_data["code"] == "т" && nuclids_exist_t) { nuclid_activity = nuclid_activity_t_real; }
                            if (nuclid_activity > 0.0f)
                            {
                                var t = -1.0f;
                                if (CodeRAO_1_MatterState == "1" && TryParseFloatExtended(nuclid_data["A_Liquid"], out A))
                                {
                                    t = T / unit_adjustment * (float)(Math.Log(nuclid_activity / (nuclid_mass_real * 1e3) / (A * 0.1f)) / 1.44f/*Math.Log(2.0)*/);
                                }
                                else if (CodeRAO_1_MatterState == "2" && TryParseFloatExtended(nuclid_data["A_Solid"], out A))
                                {
                                    t = T / unit_adjustment * (float)(Math.Log(nuclid_activity / (nuclid_mass_real * 1e3) / A) / 1.44f/*Math.Log(2.0)*/);
                                }
                                expectedPeriod = Math.Max(t, expectedPeriod);
                            }
                        }
                    }
                    if (expectedPeriod > 500.0f) { expectedValue = Math.Max(expectedValue, 3); }
                    else if (expectedPeriod >= 100.0f) { expectedValue = Math.Max(expectedValue, 2); }
                    else if (expectedPeriod >= 0.0f) { expectedValue = Math.Max(expectedValue, 1); }
                    else { expectedValue = Math.Max(expectedValue, 0); }
                    if (expectedPeriod >= 0.0f)
                    {
                        if (expectedValue.ToString("D1") != CodeRAO_6_DangerPeriod)
                        {
                            if (!soft_warning)
                            {
                                result.Add(new CheckError
                                {
                                    FormNum = "form_16",
                                    Row = forms[line].NumberInOrder_DB.ToString(),
                                    Column = "CodeRAO_DB",
                                    Value = CodeRAO_6_DangerPeriod,
                                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + $"Расчетное значение периода потенциальной опасности (в годах): {expectedPeriod} (6-ой символ кода РАО {expectedValue})."
                                });
                            }
                            else
                            {
                                result.Add(new CheckError
                                {
                                    FormNum = "form_16",
                                    Row = forms[line].NumberInOrder_DB.ToString(),
                                    Column = "CodeRAO_DB",
                                    Value = CodeRAO_6_DangerPeriod,
                                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + $"(Справочное сообщение, не обязательно к исправлению) Расчетное значение периода потенциальной опасности для приведенного полинуклидного состава (в годах): {expectedPeriod} (6-ой символ кода РАО предположительно {expectedValue})."
                                });
                            }
                        }
                    }
                }
            }
        }
        #endregion

        #region symbol 7

        if (!CodeRAO_7_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_7_RecycleMethod,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 7-го символа кода РАО."
            });
        }
        else
        {
            if (CodeRAO_1_MatterState == "1" && CodeRAO_7_RecycleMethod != "0")
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "CodeRAO_DB",
                    Value = CodeRAO_7_RecycleMethod,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для жидких РАО 7-ой символ кода РАО не может быть равным 1, 2, 3, 4, 9."
                });
            }
            else if (forms[line].OperationCode_DB == "56")
            {
                Dictionary<string, string[]> validRecycles = new() {
                    { "0", new[]
                        {
                        "11","12","13","14","15","16","17",
                        "19",
                        "21","22","23","24",
                        "29",
                        "51","52","53","54","55",
                        "61",
                        "72","73","74",
                        "99"
                        }
                    },
                    { "1", new[]
                        {
                        "31","32","39"
                        }
                    },
                    { "2", new[]
                        {
                        "41"
                        }
                    },
                    { "3", new[]
                        {
                        "42","71"
                        }
                    },
                    { "4", new[]
                        {
                        "43"
                        }
                    },
                    { "9", new[]
                        {
                        "49"
                        }
                    }
                };
                if (!validRecycles[CodeRAO_7_RecycleMethod].Contains(forms[line].RefineOrSortRAOCode_DB))
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_16",
                        Row = forms[line].NumberInOrder_DB.ToString(),
                        Column = "CodeRAO_DB",
                        Value = CodeRAO_7_RecycleMethod,
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "7-ой символ кода РАО не соответствует коду переработки/сортировки, указанному в графе 22."
                    });
                }
            }
        }
        #endregion

        #region symbol 8

        if (!CodeRAO_8_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_8_RAOClass,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 8-го символа кода РАО."
            });
        }
        else
        {
            var validValues = new string[] { "0", "7", "9" };
            if (!validValues.Contains(CodeRAO_8_RAOClass))
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "CodeRAO_DB",
                    Value = CodeRAO_8_RAOClass,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Сведения о кондиционированных отходах необходимо представлять в формах 1.7 и 1.8."
                });
            }
            else if (CodeRAO_8_RAOClass == "7")
            {
                var storageCode_1 = forms[line].StoragePlaceCode_DB.Substring(0, 1);
                var storageCodeValids = new string[] { "3", "5", "6" };
                if (!(storageCodeValids.Contains(storageCode_1) && forms[line].StatusRAO_DB == "1"))
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_16",
                        Row = forms[line].NumberInOrder_DB.ToString(),
                        Column = "CodeRAO_DB",
                        Value = CodeRAO_8_RAOClass,
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Особые РАО могут быть размещены только в ПРОРАО либо ПКОРАО."
                    });
                }
            }
            else if (CodeRAO_8_RAOClass == "9")
            {
                var storageCode_1 = forms[line].StoragePlaceCode_DB.Substring(0, 1);
                if (!(storageCode_1 == "2" && forms[line].StatusRAO_DB == "1"))
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_16",
                        Row = forms[line].NumberInOrder_DB.ToString(),
                        Column = "CodeRAO_DB",
                        Value = CodeRAO_8_RAOClass,
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Идентификатор 9 используется только для тех РАО, по которым решение об отнесении к особым или удаляемым отложено в ходе проведения первичной регистрации."
                    });
                }
            }
        }
        #endregion

        #region symbols 9-10

        if (!CodeRAO_910_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_910_TypeCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 9-10 символов кода РАО."
            });
        }
        else
        {
            if (CodeRAO_910_TypeCode == "94")
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "CodeRAO_DB",
                    Value = CodeRAO_910_TypeCode,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Сведения о РАО, подготовленных для передачи национальному оператору, предоставляются с форме 1.8."
                });
            }
            else
            {
                var requiresNote = new string[] { "19", "29", "39", "59", "69", "78", "79", "89", "99" };
                if (requiresNote.Contains(CodeRAO_910_TypeCode) && !noteExists)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_16",
                        Row = forms[line].NumberInOrder_DB.ToString(),
                        Column = "CodeRAO_DB",
                        Value = CodeRAO_910_TypeCode,
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить примечание к коду типа РАО."
                    });
                }
            }
        }
        #endregion

        #region symbol 11

        if (!CodeRAO_11_Valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "CodeRAO_DB",
                Value = CodeRAO_11_Flammability,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Недопустимое значение 11-го символа кода РАО."
            });
        }
        else
        {

        }
        #endregion

        return result;
    }

    #endregion

    #region Check005_11
    private static List<CheckError> Check_005_11(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "11", "12", "13", "14", "28", "38", "41" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var repOKPO_List = forms10.Select(x => x.Okpo_DB).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        var valid = repOKPO_List.Contains(StatusRAO_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При данном коде операции, статус РАО должен быть равен коду ОКПО отчитывающейся организации или юридического лица (РАО, образовавшиеся после 15.07.2011, находятся в собственности организации, в результате деятельности которой они образовались)."
            });
        }
        return result;
    }

    #endregion

    #region Check005_76
    private static List<CheckError> Check_005_76(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "76" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var valid = StatusRAO_DB is "6";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность статуса РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check005_Other

    private static List<CheckError> Check_005_Other(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "10", "18", "21", "25", "27", "29", "31", "35", "36", "37", "39", "44", "45", "48", "49", "51", "52", "56", "57", "59", "68", "71", "74", "75" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var applicableRAOStatuses = new string[] { "1", "2", "3", "4", "6", "9" };
        var StatusRAO_DB = forms[line].StatusRAO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(StatusRAO_DB) || applicableRAOStatuses.Contains(StatusRAO_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "StatusRAO_DB",
                Value = StatusRAO_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполнение графы 5 не соответствует требованиям приказа Госкорпорации \"Росатом\" от 07.12.2020 № 1/13-НПА"
            });
        }
        return result;
    }

    #endregion

    #region Check006
    private static List<CheckError> Check_006(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var Volume_DB = forms[line].Volume_DB;
        var valid = TryParseFloatExtended(Volume_DB, out var value) && value > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "Volume_DB",
                Value = Volume_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения об объеме РАО, если представляемые данные являются расчетными, то соответствующие значения указываются в круглых скобках"
            });
        }
        return result;
    }

    #endregion

    #region Check007
    private static List<CheckError> Check_007(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var Mass_DB = forms[line].Mass_DB;
        var Volume_DB = forms[line].Volume_DB;
        float Mass_Real;
        float Volume_Real;
        var valid = TryParseFloatExtended(Mass_DB, out Mass_Real) && Mass_Real > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "Mass_DB",
                Value = Mass_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения о массе РАО, если представляемые данные являются расчетными, то соответствующие значения указываются в круглых скобках"
            });
            return result;
        }
        if (!float.TryParse(ConvertStringToExponential(Volume_DB), out Volume_Real)) return result;
        if (Volume_Real == 0.0) return result;
        var Density_Real = Mass_Real / Volume_Real;
        if (Density_Real > 21.0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "Mass_DB",
                Value = Density_Real.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение массы и объема. Расчетное значение плотности слишком большое"
            });
        }
        else if (Density_Real < 0.005)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "Mass_DB",
                Value = Density_Real.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение массы и объема. Расчетное значение плотности слишком маленькое"
            });
        }
        return result;
    }

    #endregion

    #region Check008
    private static List<CheckError> Check_008(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var CodeRAO_DB = forms[line].CodeRAO_DB.Trim();
        var QuantityOZIII_DB = forms[line].QuantityOZIII_DB.Trim();
        var QuantityOZIII_exists = int.TryParse(QuantityOZIII_DB, out var QuantityOZIII_Real);
        var RAOTypes_1 = new string[] { "81", "82", "84", "85", "86", "87", "88", "89" };
        var RAOTypes_2 = new string[] { "99" };
        if (CodeRAO_DB.Length < 10) return result;
        var TypeRAO_DB = CodeRAO_DB.Substring(8, 2);
        var StateRAO_DB = CodeRAO_DB.Substring(0, 1);
        if (RAOTypes_1.Contains(TypeRAO_DB))
        {
            if (!QuantityOZIII_exists)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "QuantityOZIII_DB",
                    Value = QuantityOZIII_DB,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения о количестве ОЗИИИ"
                });
            }
        }
        else if (RAOTypes_2.Contains(TypeRAO_DB))
        {
            if (!QuantityOZIII_exists && QuantityOZIII_DB != "-")
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "QuantityOZIII_DB",
                    Value = QuantityOZIII_DB,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо заполнить сведения о количестве ОЗИИИ"
                });
            }
        }
        else
        {
            if (QuantityOZIII_DB != "-" || (QuantityOZIII_exists && StateRAO_DB == "2"))
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "QuantityOZIII_DB",
                    Value = QuantityOZIII_DB,
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа заполняется только для ОЗИИИ. Поставьте прочерк."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check009
    private static List<CheckError> Check_009(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].MainRadionuclids_DB;
        if (string.IsNullOrWhiteSpace(radionuclids))
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните графу 9 «Основные радионуклиды»"
            });
            return result;
        }
        var CodeRAO_DB = forms[line].CodeRAO_DB.Trim();
        if (CodeRAO_DB.Length < 3) return result;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(',', ';')
            .Split(';');
        if (radArray.Length == 1 && string.Equals(radArray[0], "-"))
        {
            if (CodeRAO_DB.Substring(2, 1) != "0")
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "Radionuclids_DB",
                    Value = Convert.ToString(radionuclids),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните графу 9 «Основные радионуклиды»"
                });
                return result;
            }
        }
        else
        {
            var unknown_nuclids = new List<string>();
            unknown_nuclids.AddRange(radArray.Where(rad => !R.Any(phEntry => phEntry["name"].Replace(" ", string.Empty) == rad)));
            var valid = unknown_nuclids.Count == 0;
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "Radionuclids_DB",
                    Value = Convert.ToString(radionuclids),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу. Наименование радионуклида указывается названием химического элемента на русском языке с указанием через дефис массового числа изотопа"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check010
    private static List<CheckError> Check_010(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].TritiumActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "т"))) return result;
        var activityReal = 1.0;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out activityReal))
        {
            if (activityReal is <= 10e+01 or > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "TritiumActivity_DB",
                    Value = Convert.ToString(forms[line].TritiumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 10"
                });
            }
            else if (activityReal > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "TritiumActivity_DB",
                    Value = Convert.ToString(forms[line].TritiumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 10. Указанная суммарная активность превышает предельное значение"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "TritiumActivity_DB",
                    Value = Convert.ToString(forms[line].TritiumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 10"
                });
            }
        }
        return result;
    }
    #endregion

    #region Check011
    private static List<CheckError> Check_011(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].BetaGammaActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "б"))) return result;
        var activityReal = 1.0;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out activityReal))
        {
            if (activityReal is <= 10e+01 or > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "BetaGammaActivity_DB",
                    Value = Convert.ToString(forms[line].BetaGammaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 11"
                });
            }
            else if (activityReal > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "BetaGammaActivity_DB",
                    Value = Convert.ToString(forms[line].BetaGammaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 11. Указанная суммарная активность превышает предельное значение"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "BetaGammaActivity_DB",
                    Value = Convert.ToString(forms[line].BetaGammaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 11"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check012
    private static List<CheckError> Check_012(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].AlphaActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "а"))) return result;
        var activityReal = 1.0;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out activityReal))
        {
            if (activityReal is <= 10e+01 or > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "AlphaActivity_DB",
                    Value = Convert.ToString(forms[line].AlphaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 12"
                });
            }
            else if (activityReal > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "AlphaActivity_DB",
                    Value = Convert.ToString(forms[line].AlphaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 12. Указанная суммарная активность превышает предельное значение"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "AlphaActivity_DB",
                    Value = Convert.ToString(forms[line].AlphaActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 12"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check013
    private static List<CheckError> Check_013(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].TransuraniumActivity_DB;
        var radionuclids = forms[line].MainRadionuclids_DB;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (!radArray.Any(rad => R.Any(phEntry => phEntry["name"] == rad && phEntry["code"] == "у")))
            return result;
        if (!double.TryParse(ConvertStringToExponential(activity),
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out var activityReal))
        {
            if (activityReal is <= 10e+01)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "TransuraniumActivity_DB",
                    Value = Convert.ToString(forms[line].TransuraniumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 13"
                });
            }
            else if (activityReal > 10e+20)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "TransuraniumActivity_DB",
                    Value = Convert.ToString(forms[line].TransuraniumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение суммарной активности в графе 13. Указанная суммарная активность превышает предельное значение"
                });
            }
            else
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "TransuraniumActivity_DB",
                    Value = Convert.ToString(forms[line].TransuraniumActivity_DB),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для указанного в графе 9 радионуклидного состава должна быть приведена активность в графе 13"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check014
    private static List<CheckError> Check_014(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var activityDate = forms[line].ActivityMeasurementDate_DB;
        var operationDate = forms[line].OperationDate_DB;
        DateTime activityDateReal = DateTime.Now;
        DateTime operationDateReal = DateTime.UnixEpoch;
        var valid = DateTime.TryParse(activityDate, out activityDateReal) && DateTime.TryParse(operationDate, out operationDateReal);
        if (valid)
        {
            valid = activityDateReal <= operationDateReal;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "ActivityMeasurementDate_DB",
                Value = activityDate,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность указанной даты измерения активности. Дата не может быть позже даты операции"
            });
        }
        return result;
    }

    #endregion

    #region Check015
    private static List<CheckError> Check_015(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        byte?[] validDocument = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 19 };
        var DocumentVid_DB = forms[line].DocumentVid_DB;
        var valid = validDocument.Contains(DocumentVid_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "DocumentVid_DB",
                Value = DocumentVid_DB.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте правильность заполнения графы 15"
            });
        }
        return result;
    }

    #endregion

    #region Check016
    private static List<CheckError> Check_016(List<Form16> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var documentNumberDB = forms[line].DocumentNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(documentNumberDB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "DocumentNumber_DB",
                Value = Convert.ToString(documentNumberDB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check017
    private static List<CheckError> Check_017(List<Form16> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        var documentDate = forms[line].DocumentDate_DB;
        DateTime pMid;
        DateTime pEnd;
        var valid = true;
        if (operationCode == "41")
        {
            valid = DateTime.TryParse(documentDate, out pMid)
                    && DateTime.TryParse(operationDate, out var pOper)
                    && pMid.Date == pOper.Date;
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "DocumentDate_DB",
                    Value = Convert.ToString(documentDate),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа должна соответствовать дате операции"
                });
            }
        }
        else if (operationCode == "10")
        {
            valid = DateTime.TryParse(rep.StartPeriod_DB, out var pStart)
                    && DateTime.TryParse(rep.EndPeriod_DB, out pEnd)
                    && DateTime.TryParse(documentDate, out pMid)
                    && pMid >= pStart && pMid <= pEnd;
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "DocumentDate_DB",
                    Value = Convert.ToString(documentDate),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа выходит за границы периода"
                });
            }
        }
        else
        {
            valid = DateTime.TryParse(documentDate, out pMid)
                    && DateTime.TryParse(operationDate, out var pOper)
                    && pMid.Date <= pOper.Date;
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "DocumentDate_DB",
                    Value = Convert.ToString(documentDate),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа не может быть позже даты операции"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check018_10

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_018_10(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10", "11", "12", "13", "14", "16", "18", "41", "42", "43", "44", "45", "48", "49", "51", "52", "56", "57", "59", "68", "71", "72", "73", "74", "75", "76", "97", "98" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = providerOrRecieverOKPO == repOKPO;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для выбранного кода операции указывается код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check018_21

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_018_21(List<Form16> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "21", "25", "26", "27", "28", "29", "31", "35", "36", "37", "38", "39", "99" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Значение может состоять только из 8 или 14 символов"
            });
        }
        valid = providerOrRecieverOKPO != repOKPO;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для выбранного кода операции указывается код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check018_22

    //Код ОКПО поставщика/получателя состоит из 8/14 чисел или "минобороны" (колонка 18)
    private static List<CheckError> Check_018_22(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (operationCode is not ("22" or "32")) return result;
        var valid = okpoRegex.IsMatch(providerOrRecieverOKPO)
                    || providerOrRecieverOKPO.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу. Следует указать код ОКПО контрагента, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check019_01

    //При определенных кодах операции, код ОКПО перевозчика равен "-"
    private static List<CheckError> Check_019_01(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = {
            "01", "10", "11","12","13","14","16", "18", "41", "42","43", "44", "45", "48", "49", "51", "52","56", "57", "59", "71", "72", "73", "74", "75",
            "76", "97", "98", "99"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOKPO = forms[line].TransporterOKPO_DB.Trim();
        var valid = transporterOKPO is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При выбранном коде операции транспортирование не производится, в графе 19 должен стоять прочерк."
            });
        }
        return result;
    }

    #endregion

    #region Check019_21

    //При определенных кодах операции, код ОКПО перевозчика равен 8/14 цифр
    private static List<CheckError> Check_019_21(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21","25","26","27","28","29","31","35","36","37","38","39"
        };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        var valid = okpoRegex.IsMatch(transporterOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Значение может состоять только из 8 или 14 символов"
            });
        }
        return result;
    }

    #endregion

    #region Check019_22

    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны"
    private static List<CheckError> Check_019_22(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "22", "32" };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = okpoRegex.IsMatch(transporterOKPO)
                    || transporterOKPO.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо указать код ОКПО организации перевозчика, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check020
    private static List<CheckError> Check_020(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var StoragePlaceName_DB = forms[line].StoragePlaceName_DB;
        var valid = !string.IsNullOrWhiteSpace(StoragePlaceName_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "StoragePlaceName_DB",
                Value = Convert.ToString(StoragePlaceName_DB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа 20 должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check021
    private static List<CheckError> Check_021(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var StoragePlaceCode_DB = forms[line].StoragePlaceCode_DB;
        var valid = !string.IsNullOrWhiteSpace(StoragePlaceCode_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "StoragePlaceCode_DB",
                Value = Convert.ToString(StoragePlaceCode_DB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа 21 должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check022_44
    private static List<CheckError> Check_022_44(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "44" };
        var applicableValues = new string[] {
            "11","12","13","14","15","16","17","19",
            "21","22","23","24",               "29",
            "31","32",                         "39",
            "41","42","43",                    "49",
            "51","52","53","54","55",
            "61","62","63",
            "71","72","73","74",               "79",
                                               "99"
        };
        var operationCode = forms[line].OperationCode_DB.Trim();
        var RefineOrSortRAOCode_DB = forms[line].RefineOrSortRAOCode_DB.Trim();
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = applicableValues.Contains(RefineOrSortRAOCode_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + $"Для кода операции {operationCode} код переработки/сортировки {RefineOrSortRAOCode_DB} недопустим."
            });
        }
        return result;
    }

    #endregion

    #region Check022_45
    private static List<CheckError> Check_022_45(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "45", "57" };
        var applicableValues = new string[] { "-", "74" };
        var operationCode = forms[line].OperationCode_DB.Trim();
        var RefineOrSortRAOCode_DB = forms[line].RefineOrSortRAOCode_DB.Trim();
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = applicableValues.Contains(RefineOrSortRAOCode_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Коду операции упаковка/переупаковка не соответствует код переработки/сортировки"
            });
        }
        return result;
    }

    #endregion

    #region Check022_49
    private static List<CheckError> Check_022_49(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "49", "59" };
        var applicableRefineCodes = new string[] { "52", "72", "74", "-" };
        var operationCode = forms[line].OperationCode_DB.Trim();
        var RefineOrSortRAOCode_DB = forms[line].RefineOrSortRAOCode_DB.Trim();
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = applicableRefineCodes.Contains(RefineOrSortRAOCode_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Коду операции сортировка соответствуют коды сортировки 52, 72, 74"
            });
        }
        return result;
    }

    #endregion

    #region Check022_10
    private static List<CheckError> Check_022_10(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var applicableOperationCodes = new string[] { "10","11","12","13","14","16","18","21","22","25","26","27","28","29","31","32","35",
            "36","37","38","39","42","43","48","51","52","63","64","68","71","72","73","74","75",
            "76","97","98","99" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = forms[line].RefineOrSortRAOCode_DB.Trim() == "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При данном коде операции для кода переработки/сортировки используется символ «-»"
            });
        }
        return result;
    }

    #endregion

    #region Check022_22_44
    private static List<CheckError> Check_022_22_44(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].RefineOrSortRAOCode_DB != "22") return result;
        var CodeRAO_DB = forms[line].CodeRAO_DB;
        if (CodeRAO_DB.Length < 11) return result;
        if (forms[line].OperationCode_DB != "44") return result;
        var valid = CodeRAO_DB.Substring(10, 1) == "1";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "РАО направлены на установку сжигания. Проверьте значение 11 символа кода РАО"
            });
        }
        return result;
    }

    #endregion

    #region Check022_22_56
    private static List<CheckError> Check_022_22_56(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].RefineOrSortRAOCode_DB != "22") return result;
        var CodeRAO_DB = forms[line].CodeRAO_DB;
        if (CodeRAO_DB.Length < 10) return result;
        if (forms[line].OperationCode_DB != "56") return result;
        var valid = (CodeRAO_DB.Substring(0, 1) is "2" && CodeRAO_DB.Substring(8, 2) is "66" or "74") || (CodeRAO_DB.Substring(0, 1) is "1" && CodeRAO_DB.Substring(8, 2) is "14");
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = forms[line].RefineOrSortRAOCode_DB,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "РАО направлены на переработку на установке сжигания. Проверьте значение кода типа РАО в коде РАО."
            });
        }
        return result;
    }

    #endregion

    #region Check023
    private static List<CheckError> Check_023(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackName_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "PackName_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните сведения об упаковке РАО. Если РАО размещены без упаковки, то в графе 23 указывается «без упаковки»."
            });
        }
        return result;
    }

    #endregion

    #region Check024
    private static List<CheckError> Check_024(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackType_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "PackType_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните сведения в графе 24. В случае, если тип отсутствует, укажите символ «-» без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check025
    private static List<CheckError> Check_025(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "PackNumber_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните сведения о заводском номере упаковки. Если заводской номер отсутствует необходимо привести в круглых скобках номер, присвоенный в организации."
            });
        }
        return result;
    }

    #endregion

    #region Check026
    private static List<CheckError> Check_026(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].Subsidy_DB;
        if (string.IsNullOrWhiteSpace(field_value) || field_value.Trim() == "-") return result;
        if (float.TryParse(field_value, out var value_real))
        {
            if (value_real < 0 || value_real > 100)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_16",
                    Row = forms[line].NumberInOrder_DB.ToString(),
                    Column = "Subsidy_DB",
                    Value = Convert.ToString(field_value),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Проверьте значение субсидии."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check027

    private static List<CheckError> Check_027(List<Form16> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].FcpNumber_DB;
        var valid = string.IsNullOrWhiteSpace(field_value) || field_value.Trim() == "-" || TryParseFloatExtended(field_value.Trim(), out var val);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = forms[line].NumberInOrder_DB.ToString(),
                Column = "FcpNumber_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check028

    //Наличие строк дубликатов
    private static List<CheckError> Check_028(List<Form16> forms)
    {
        List<CheckError> result = new();
        HashSet<int> duplicatesLinesSet = new();
        var comparator = new CustomNullStringWithTrimComparer();
        for (var i = 0; i < forms.Count; i++)
        {
            var currentForm = forms[i];
            for (var j = i + 1; j < forms.Count; j++)
            {
                var formToCompare = forms[j];
                var isDuplicate = comparator.Compare(formToCompare.OperationCode_DB, currentForm.OperationCode_DB) == 0
                                  && comparator.Compare(formToCompare.OperationDate_DB, currentForm.OperationDate_DB) == 0
                                  && comparator.Compare(formToCompare.CodeRAO_DB, currentForm.CodeRAO_DB) == 0
                                  && comparator.Compare(formToCompare.StatusRAO_DB, currentForm.StatusRAO_DB) == 0
                                  && comparator.Compare(formToCompare.Volume_DB, currentForm.Volume_DB) == 0
                                  && comparator.Compare(formToCompare.Mass_DB, currentForm.Mass_DB) == 0
                                  && formToCompare.QuantityOZIII_DB == currentForm.QuantityOZIII_DB
                                  && comparator.Compare(formToCompare.MainRadionuclids_DB, currentForm.MainRadionuclids_DB) == 0
                                  && comparator.Compare(formToCompare.TritiumActivity_DB, currentForm.TritiumActivity_DB) == 0
                                  && comparator.Compare(formToCompare.BetaGammaActivity_DB, currentForm.BetaGammaActivity_DB) == 0
                                  && comparator.Compare(formToCompare.AlphaActivity_DB, currentForm.AlphaActivity_DB) == 0
                                  && comparator.Compare(formToCompare.TransuraniumActivity_DB, currentForm.TransuraniumActivity_DB) == 0
                                  && comparator.Compare(formToCompare.ActivityMeasurementDate_DB, currentForm.ActivityMeasurementDate_DB) == 0
                                  && formToCompare.DocumentVid_DB == currentForm.DocumentVid_DB
                                  && comparator.Compare(formToCompare.DocumentNumber_DB, currentForm.DocumentNumber_DB) == 0
                                  && comparator.Compare(formToCompare.DocumentDate_DB, currentForm.DocumentDate_DB) == 0
                                  && comparator.Compare(formToCompare.ProviderOrRecieverOKPO_DB, currentForm.ProviderOrRecieverOKPO_DB) == 0;
                if (!isDuplicate) continue;
                duplicatesLinesSet.Add(i + 1);
                duplicatesLinesSet.Add(j + 1);
            }
        }
        var duplicateLines = string.Join(", ", duplicatesLinesSet.Order());
        if (duplicatesLinesSet.Count != 0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = duplicateLines,
                Column = "2 - 18",
                Value = "",
                Message = $"Данные граф 2-18 в строках {duplicateLines} продублированы. Следует проверить правильность предоставления данных."
            });
        }
        return result;
    }

    #endregion

    #region Check029
    //overdue calculations
    private static List<CheckError> Check_029(List<Form16> forms, Report rep)
    {
        List<CheckError> result = new();
        List<string> overdueSet = new();
        var EndPeriod_DB = rep.EndPeriod_DB;
        if (DateOnly.TryParse(EndPeriod_DB, out var date_end))
        {
            foreach (Form16 form in forms)
            {
                var operationCode = form.OperationCode_DB;
                var operationDate = form.OperationDate_DB;
                var documentDate = form.DocumentDate_DB;
                if (OverduePeriods_RAO.TryGetValue(operationCode, out var days) && DateOnly.TryParse(operationCode == "10" ? documentDate : operationDate, out var date_mid))
                {
                    if (WorkdaysBetweenDates(date_mid, date_end) > days)
                    {
                        overdueSet.Add($"[{operationCode} @ {date_mid} ({WorkdaysBetweenDates(date_mid, date_end) - days})]");
                    }
                }
            }
        }
        if (overdueSet.Count > 0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_16",
                Row = string.Join(", ", overdueSet),
                Column = "-",
                Value = "-",
                Message = $"Указанные операции просрочены."
            });
        }
        return result;
    }
    #endregion

    #endregion
}
