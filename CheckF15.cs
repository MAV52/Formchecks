using Models.CheckForm;
using Models.Collections;
using Models.Forms;
using Models.Forms.Form1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Client_App.Commands.SyncCommands.CheckForm;

public class CheckF15 : CheckBase
{
    #region Properties

    private static readonly string[] OperationCode_DB_Valids =
    {
        "01","10","14","21","22","25","26",
        "27","28","29","31","32","35","36",
        "37","38","39","43","44","45","49",
        "51","52","57","59","63","64","71",
        "72","73","74","75","76","84","88",
        "97","98","99",
        "41"
    };

    private static readonly Dictionary<string, string> GraphsList = new()
    {
        { "NumberInOrder_DB", "01 - № п/п" },
        { "OperationCode_DB", "02 - Код операции" },
        { "OperationDate_DB", "03 - Дата операции" },
        { "PassportNumber_DB", "04 - Номер паспорта (сертификата)" },
        { "Type_DB", "05 - Тип ЗРИ" },
        { "Radionuclids_DB", "06 - Радионуклиды" },
        { "FactoryNumber_DB", "07 - Номер" },
        { "Quantity_DB", "08 - Количество, шт." },
        { "Activity_DB", "09 - Суммарная активность, Бк" },
        { "CreationDate_DB", "10 - Дата выпуска" },
        { "StatusRAO_DB", "11 - Статус РАО" },
        { "DocumentVid_DB", "12 - Вид документа" },
        { "DocumentNumber_DB", "13 - Номер документа" },
        { "DocumentDate_DB", "14 - Дата документа" },
        { "ProviderOrRecieverOKPO_DB", "15 - Код ОКПО поставщика или получателя" },
        { "TransporterOKPO_DB", "16 - Код ОКПО перевозчика" },
        { "PackName_DB", "17 - Наименование прибора, УКТ, упаковки" },
        { "PackType_DB", "18 - Тип прибора, УКТ, упаковки" },
        { "PackNumber_DB", "19 - Заводской номер прибора, УКТ, упаковки" },
        { "StoragePlaceName_DB", "20 - Наименование пункта хранения" },
        { "StoragePlaceCode_DB", "21 - Код пункта хранения" },
        { "RefineOrSortRAOCode_DB", "22 - Код переработки/сортировки РАО" },
        { "Subsidy_DB", "23 - Субсидия, %" },
        { "FcpNumber_DB", "24 - Номер мероприятия ФЦП" },
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
        if (D.Count == 0)
        {
#if DEBUG
            D_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\")), "data", "Spravochniki", "D.xlsx"));
#else
            D_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"D.xlsx"));
#endif
        }
        if (holidays_specific.Count == 0)
        {
#if DEBUG
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\..\\")), "data", "Spravochniki", "Holidays.xlsx"));
#else
            Holidays_Populate_From_File(Path.Combine(Path.GetFullPath(AppContext.BaseDirectory), "data", "Spravochniki", $"Holidays.xlsx"));
#endif
        }
        foreach (var key in rep.Rows15)
        {
            var form = (Form15)key;
            var formsList = rep.Rows11.ToList<Form15>();
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_001(formsList, currentFormLine));
            errorList.AddRange(Check_002(formsList, currentFormLine));
            errorList.AddRange(Check_002_10(formsList, currentFormLine));
            errorList.AddRange(Check_002_29(formsList, notes, currentFormLine));
            errorList.AddRange(Check_002_21(formsList, currentFormLine));
            errorList.AddRange(Check_002_37(formsList, currentFormLine));
            errorList.AddRange(Check_002_41(formsList, currentFormLine));
            errorList.AddRange(Check_002_51(formsList, currentFormLine));
            errorList.AddRange(Check_002_52(formsList, currentFormLine));
            errorList.AddRange(Check_002_57(formsList, currentFormLine));
            errorList.AddRange(Check_002_84(formsList, currentFormLine));
            errorList.AddRange(Check_002_71(formsList, currentFormLine));
            errorList.AddRange(Check_003_03(formsList, rep, currentFormLine));
            errorList.AddRange(Check_003_15(formsList, rep, currentFormLine));
            errorList.AddRange(Check_004(formsList, currentFormLine));
            errorList.AddRange(Check_005(formsList, currentFormLine));
            errorList.AddRange(Check_006_a(formsList, currentFormLine));
            errorList.AddRange(Check_006_b(formsList, currentFormLine));
            errorList.AddRange(Check_007_a(formsList, currentFormLine));
            errorList.AddRange(Check_007_b(formsList, currentFormLine));
            errorList.AddRange(Check_008(formsList, currentFormLine));
            errorList.AddRange(Check_009(formsList, currentFormLine));
            errorList.AddRange(Check_010(formsList, currentFormLine));
            errorList.AddRange(Check_011_OKPO(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_011_76(formsList, currentFormLine));
            errorList.AddRange(Check_011_41(formsList, currentFormLine));
            errorList.AddRange(Check_011_misc(formsList, currentFormLine));
            errorList.AddRange(Check_012(formsList, notes, currentFormLine));
            errorList.AddRange(Check_013(formsList, currentFormLine));
            errorList.AddRange(Check014(formsList, rep, currentFormLine));
            errorList.AddRange(Check_015_01(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_015_21(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_015_22(formsList, currentFormLine));
            errorList.AddRange(Check_015_84(formsList, notes, currentFormLine));
            errorList.AddRange(Check_016_01(formsList, currentFormLine));
            errorList.AddRange(Check_016_21(formsList, currentFormLine));
            errorList.AddRange(Check_016_22(formsList, currentFormLine));
            errorList.AddRange(Check_017(formsList, currentFormLine));
            errorList.AddRange(Check_018(formsList, currentFormLine));
            errorList.AddRange(Check_019(formsList, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, currentFormLine));
            errorList.AddRange(Check_022(formsList, currentFormLine));
            errorList.AddRange(Check_023(formsList, currentFormLine));
            errorList.AddRange(Check_024(formsList, currentFormLine));

            currentFormLine++;
        }

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

    private static List<CheckError> Check_001(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].NumberInOrder_DB != line + 1)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "NumberInOrder_DB",
                Value = forms[line].NumberInOrder_DB.ToString(),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Номера строк должны располагаться по порядку, без пропусков или дублирования номеров"
            });
        }
        return result;
    }

    #endregion

    #region Check002

    private static List<CheckError> Check_002(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var valid = operationCode != null && OperationCode_DB_Valids.Contains(operationCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Код операции не может быть использован в форме 1.1."
            });
        }
        return result;
    }

    #endregion

    #region Check002_10

    private static List<CheckError> Check_002_10(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "10" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Сведения, представленные в инвентаризации, не соответствуют СНК"
            });
        }
        return result;
    }

    #endregion

    #region Check002_29

    private static List<CheckError> Check_002_29(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "29,39,49,59,97,98,99" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        const byte graphNumber = 2;
        var valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо дать пояснение об осуществленной операции."
            });
        }
        return result;
    }

    #endregion

    #region Check002_21

    private static List<CheckError> Check_002_21(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "21,22,25,26,27,28,29,42,43,44,45,49,51,71,72,84,98" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Учетной единицы с такими параметрами нет в организации. Проверьте правильность указываемых сведений для ОЗРИ"
            });
        }
        return result;
    }

    #endregion

    #region Check002_37

    private static List<CheckError> Check_002_37(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "37" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В отчетах не найдена строка об осуществлении передачи учетной единицы. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_41
    //эта проверка по идее никогда не выдаст ошибку, т.к. проверке все равно, появился ли 41 код вручную или автоматически.
    private static List<CheckError> Check_002_41(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "41" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = true;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + ""
            });
        }
        return result;
    }

    #endregion

    #region Check002_51
    private static List<CheckError> Check_002_51(List<Form15> forms, int line)
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + ""
            });
        }
        return result;
    }

    #endregion

    #region Check002_52
    private static List<CheckError> Check_002_52(List<Form15> forms, int line)
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В отчетах не найдена строка об изъятии РАО из пункта хранения. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_57
    private static List<CheckError> Check_002_57(List<Form15> forms, int line)
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "В отчетах не найдена строка снятии учетной единицы для упаковки/переупаковки. Проверьте правильность выбранного кода операции"
            });
        }
        return result;
    }

    #endregion

    #region Check002_84
    private static List<CheckError> Check_002_84(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "57" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = forms[line].StatusRAO_DB.Trim() == "7";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для ОЗРИ, которые возвращаются в страну поставщика код статус РАО – 7"
            });
        }
        return result;
    }

    #endregion

    #region Check002_71
    //Справочная "ошибка" - т.е. даже не ошибка.
    private static List<CheckError> Check_002_71(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new string[] { "71", "72", "73", "74", "75", "76" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = false;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "К отчету необходимо приложить скан-копию документа характеризующего операцию"
            });
        }
        return result;
    }

    #endregion

    #region Check003_03
    private static List<CheckError> Check_003_03(List<Form15> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] nonApplicableOperationCodes = { "01", "10" };
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(operationDate),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата операции не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion


    #region Check003_15
    private static List<CheckError> Check_003_15(List<Form15> forms, Report rep, int line)
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(documentDate),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата документа не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion

    #region Check004
    private static List<CheckError> Check_004(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var passportNumber = forms[line].PassportNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(passportNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PassportNumber_DB",
                Value = Convert.ToString(passportNumber),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check005
    private static List<CheckError> Check_005(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var typeDB = forms[line].Type_DB;
        var valid = !string.IsNullOrWhiteSpace(typeDB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Type_DB",
                Value = Convert.ToString(typeDB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check006_a

    //У радионуклидов в качестве разделителя использовать ;
    private static List<CheckError> Check_006_a(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].Radionuclids_DB;
        if (radionuclids.Contains(',') || radionuclids.Contains('+'))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу. Радионуклиды должны быть разделены точкой с запятой."
            });
        }
        return result;
    }

    #endregion

    #region Check006_b

    //Все радионуклиды есть в справочнике
    private static List<CheckError> Check_006_b(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].Radionuclids_DB;
        if (string.IsNullOrWhiteSpace(radionuclids)) return result;
        var radArray = radionuclids.Replace(" ", string.Empty).ToLower()
            .Replace(" ", "")
            .Replace(',', ';')
            .Split(';');
        if (radArray.Length == 1 && string.Equals(radArray[0], "-")) return result;
        var valid = radArray.All(rad => R.Any(phEntry => phEntry["name"] == rad));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу. Радионуклид отсутствует в справочнике."
            });
        }
        return result;
    }

    #endregion

    #region Check007_a

    //У номеров в качестве разделителя использовать ;
    private static List<CheckError> Check_007_a(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var numbers = forms[line].FactoryNumber_DB;
        if (numbers.Contains(',') || numbers.Contains('+'))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "FactoryNumber_DB",
                Value = Convert.ToString(numbers),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу. Номера должны быть разделены точкой с запятой."
            });
        }
        return result;
    }

    #endregion

    #region Check007_b

    //Номер не пустая строка (колонка 7)
    private static List<CheckError> Check_007_b(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var factoryNumber = forms[line].FactoryNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(factoryNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "FactoryNumber_DB",
                Value = Convert.ToString(factoryNumber),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните сведения о номере ЗРИ, который переведен в ОЗИИИ"
            });
        }
        return result;
    }

    #endregion

    #region Check008
    private static List<CheckError> Check_008(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var quantityDB = forms[line].Quantity_DB;
        var valid = quantityDB != null;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Quantity_DB",
                Value = Convert.ToString(quantityDB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните сведения о количестве ЗРИ, переведенных в ОЗИИИ"
            });
        }
        return result;
    }

    #endregion

    #region Check009
    private static List<CheckError> Check_009(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var activity = forms[line].Activity_DB;
        if (string.IsNullOrEmpty(activity) || activity == "-") return result;
        activity = activity
            .Replace(".", ",")
            .Replace("(", "")
            .Replace(")", "");
        if (!double.TryParse(activity,
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out var activityReal)
            || activity.Contains('-')
            || activityReal is <= 0 or > 10e+20)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = Convert.ToString(forms[line].Activity_DB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Заполните сведения о суммарной активности ЗРИ, переведенных в ОЗИИИ. Оценочные сведения приводятся в круглых скобках"
            });
        }
        return result;
    }
    #endregion

    #region Check010
    private static List<CheckError> Check_010(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationDate = forms[line].OperationDate_DB;
        var creationDate = forms[line].CreationDate_DB;
        var valid = DateTime.TryParse(creationDate, out var creationDateReal)
                    && DateTime.TryParse(operationDate, out var operationDateReal)
                    && creationDateReal <= operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "CreationDate_DB",
                Value = Convert.ToString(creationDate),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Дата выпуска не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check011_OKPO
    private static List<CheckError> Check_011_OKPO(List<Form15> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "14", "28", "38" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var status = forms[line].StatusRAO_DB;
        var repOKPO = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        var valid = status == repOKPO;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = Convert.ToString(status),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для операции с кодом 14, 28 и 38 статус РАО должен быть равен коду ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check011_76
    private static List<CheckError> Check_011_76(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "76" };
        var operationCode = forms[line].OperationCode_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var status = forms[line].StatusRAO_DB;
        var valid = status == "6";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = Convert.ToString(status),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для операции с кодом 76 статус РАО должен быть равен 6."
            });
        }
        return result;
    }

    #endregion

    #region Check011_41
    private static List<CheckError> Check_011_41(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        return result;
    }

    #endregion

    #region Check011_misc
    private static List<CheckError> Check_011_misc(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        string[] nonApplicableOperationCodes = { "14", "28", "38", "76" };
        string[] validStatus = { "1", "2", "3", "4", "6", "7", "9" };
        var operationCode = forms[line].OperationCode_DB;
        if (nonApplicableOperationCodes.Contains(operationCode)) return result;
        var status = forms[line].StatusRAO_DB;
        var valid = validStatus.Contains(status);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StatusRAO_DB",
                Value = Convert.ToString(status),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Неверный код \"Статус РАО\""
            });
        }
        return result;
    }

    #endregion

    #region Check012
    private static List<CheckError> Check_012(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 12;
        byte?[] validDocumentVid = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 19 };
        var documentVid = forms[line].DocumentVid_DB;
        var valid = validDocumentVid.Contains(documentVid);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Неверный код вида документа, сопровождающего операцию."
            });
        }
        else if (documentVid == 19)
        {
            valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_15",
                    Row = (line + 1).ToString(),
                    Column = "DocumentVid_DB",
                    Value = Convert.ToString(documentVid),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "К коду вида документа 19 необходимо примечание."
                });
            }
        }
        return result;
    }

    #endregion

    #region Check013
    private static List<CheckError> Check_013(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var documentNumberDB = forms[line].DocumentNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(documentNumberDB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentNumber_DB",
                Value = Convert.ToString(documentNumberDB),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check014
    private static List<CheckError> Check014(List<Form15> forms, Report rep, int line)
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
                    FormNum = "form_15",
                    Row = (line + 1).ToString(),
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
                    FormNum = "form_15",
                    Row = (line + 1).ToString(),
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
                    FormNum = "form_15",
                    Row = (line + 1).ToString(),
                    Column = "DocumentDate_DB",
                    Value = Convert.ToString(documentDate),
                    Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Даат документа не может быть позже даты операции"
                });
            }
        }
        return result;
    }

    #endregion

    #region Check015_01

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_015_01(List<Form15> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "01", "10", "14", "41", "43", "44", "45", "49", "51", "52", "57", "59", "71", "72", "73", "74", "75", "76", "97", "98", "99" };
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для выбранного кода операции указывается код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check015_21

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_015_21(List<Form15> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "21", "25", "26", "27", "28", "29", "31", "35", "36", "37", "38", "39" };
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для выбранного кода операции указывается код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check015_22

    //Код ОКПО поставщика/получателя состоит из 8/14 чисел или "минобороны" (колонка 18)
    private static List<CheckError> Check_015_22(List<Form15> forms, int line)
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу. Следует указать код ОКПО контрагента, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check015_84

    //Код ОКПО поставщика/получателя из ОКСМ (не Россия), для определенных кодов операции, с примечанием (колонка 18)
    private static List<CheckError> Check_015_84(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 15;
        string[] applicableOperationCodes = { "84", "88" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOKPO = forms[line].ProviderOrRecieverOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == providerOrRecieverOKPO)
                    && !providerOrRecieverOKPO.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Формат ввода данных не соответствует приказу."
            });
        }
        valid = CheckNotePresence(new List<Form>(forms), notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо добавить примечание."
            });
        }
        return result;
    }

    #endregion

    #region Check016_01

    //При определенных кодах операции, код ОКПО перевозчика равен "-"
    private static List<CheckError> Check_016_01(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = {
            "01", "10", "14", "18", "41", "43", "44", "45", "48", "49", "51", "52", "57", "59", "71", "72", "73", "74", "75",
            "76", "97", "98", "99"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        var valid = transporterOKPO is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "При выбранном коде операции транспортирование не производится."
            });
        }
        return result;
    }

    #endregion

    #region Check016_21

    //При определенных кодах операции, код ОКПО перевозчика равен 8/14 цифр
    private static List<CheckError> Check_016_21(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21","25","26","27","28","29","31","35","36","37","38","39","84","85","86","88"
        };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOKPO = forms[line].TransporterOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var okpoRegex = new Regex(@"^\d{8}([0123456789_]\d{5})?$");
        if (okpoRegex.IsMatch(transporterOKPO)) return result;

        var valid = okpoRegex.IsMatch(transporterOKPO);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Значение может состоять только из 8 или 14 символов"
            });
        }
        return result;
    }

    #endregion

    #region Check016_22

    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны"
    private static List<CheckError> Check_016_22(List<Form15> forms, int line)
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
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOKPO),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Необходимо указать код ОКПО организации перевозчика, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check017
    private static List<CheckError> Check_017(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackName_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackName_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check018
    private static List<CheckError> Check_018(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackType_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackType_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check019
    private static List<CheckError> Check_019(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].PackNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackNumber_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check020
    private static List<CheckError> Check_020(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].StoragePlaceName_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceName_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check021
    private static List<CheckError> Check_021(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].StoragePlaceCode_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceCode_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check022
    private static List<CheckError> Check_022(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var sortcode = forms[line].RefineOrSortRAOCode_DB;
        var valid = !string.IsNullOrWhiteSpace(sortcode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = Convert.ToString(sortcode),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        else
        {
            var operationCode = forms[line].OperationCode_DB;
            string[] applicableOperationCodes1 = {
                "01", "10", "14", "21", "22", "25", "26", "27", "28", "29", "31", "32", "35", "36", "37", "38", "39", "43", "51",
                "52", "63", "64", "71", "72", "73", "74", "75", "76", "84", "88", "97", "98", "99"
            };
            string[] applicableOperationCodes2 = {
                "44", "45", "49", "57", "59", "98"
            };
            string[] applicableSortCodes = { "52", "72", "74" };
            if (applicableOperationCodes1.Contains(operationCode))
            {
                valid = sortcode == "-";
                if (!valid)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "RefineOrSortRAOCode_DB",
                        Value = Convert.ToString(sortcode),
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Для данного кода операции в качестве кода переработки/сортировки указывается прочерк"
                    });
                }
            }
            else if (applicableOperationCodes2.Contains(operationCode))
            {
                valid = applicableSortCodes.Contains(sortcode);
                if (!valid)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "RefineOrSortRAOCode_DB",
                        Value = Convert.ToString(sortcode),
                        Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Коду операции сортировка соответствуют коды сортировки 52, 72, 74"
                    });
                }
            }
        }
        return result;
    }

    #endregion

    #region Check023
    private static List<CheckError> Check_023(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].Subsidy_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceCode_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check024
    private static List<CheckError> Check_024(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var field_value = forms[line].FcpNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(field_value);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceCode_DB",
                Value = Convert.ToString(field_value),
                Message = $"Проверка {MethodBase.GetCurrentMethod()?.Name.Replace("Check_", "").TrimStart('0')} - " + "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #endregion
}