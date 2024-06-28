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
    #region CheckTotal

    public static List<CheckError> Check_Total(Reports reps, Report rep)
    {
        var currentFormLine = 0;
        List<CheckError> errorList = new();
        LoadDictionaries();
        var formsList = rep.Rows15.ToList<Form15>();
        errorList.AddRange(Check_001(rep));
        errorList.AddRange(Check_002(formsList, rep));
        errorList.AddRange(Check_003(formsList));
        foreach (var key in rep.Rows15)
        {
            var form = (Form15)key;
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_004(formsList, currentFormLine));
            errorList.AddRange(Check_005(formsList, currentFormLine));
            errorList.AddRange(Check_005_10(formsList, currentFormLine));
            errorList.AddRange(Check_005_29(formsList, notes, currentFormLine));
            errorList.AddRange(Check_005_21(formsList, currentFormLine));
            errorList.AddRange(Check_005_37(formsList, currentFormLine));
            errorList.AddRange(Check_005_41(formsList, currentFormLine));
            errorList.AddRange(Check_005_51(formsList, currentFormLine));
            errorList.AddRange(Check_005_52(formsList, currentFormLine));
            errorList.AddRange(Check_005_57(formsList, currentFormLine));
            errorList.AddRange(Check_005_84(formsList, currentFormLine));
            errorList.AddRange(Check_005_71(formsList, currentFormLine));
            errorList.AddRange(Check_005_9(formsList, notes, currentFormLine));
            errorList.AddRange(Check_006(formsList, rep, currentFormLine));
            errorList.AddRange(Check_006_03(formsList, rep, currentFormLine));
            errorList.AddRange(Check_007(formsList, currentFormLine));
            errorList.AddRange(Check_008(formsList, currentFormLine));
            errorList.AddRange(Check_009_a(formsList, currentFormLine));
            errorList.AddRange(Check_009_b(formsList, currentFormLine));
            errorList.AddRange(Check_010_a(formsList, currentFormLine));
            errorList.AddRange(Check_010_b(formsList, currentFormLine));
            errorList.AddRange(Check_011(formsList, currentFormLine));
            errorList.AddRange(Check_012(formsList, currentFormLine));
            errorList.AddRange(Check_013(formsList, currentFormLine));
            errorList.AddRange(Check_014(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_015(formsList, notes, currentFormLine));
            errorList.AddRange(Check_016(formsList, currentFormLine));
            errorList.AddRange(Check_017_a(formsList, currentFormLine));
            errorList.AddRange(Check_017_b(formsList, currentFormLine));
            errorList.AddRange(Check_017_c(formsList, currentFormLine));
            errorList.AddRange(Check_017_d(formsList, rep, currentFormLine));
            errorList.AddRange(Check_018_01(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_018_21(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_018_22(formsList, currentFormLine));
            errorList.AddRange(Check_018_84(formsList, notes, currentFormLine));
            errorList.AddRange(Check_019_01(formsList, currentFormLine));
            errorList.AddRange(Check_019_21(formsList, currentFormLine));
            errorList.AddRange(Check_019_22(formsList, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, currentFormLine));
            errorList.AddRange(Check_022(formsList, currentFormLine));
            errorList.AddRange(Check_023(formsList, currentFormLine));
            errorList.AddRange(Check_024(formsList, currentFormLine));
            errorList.AddRange(Check_025(formsList, notes, currentFormLine));
            errorList.AddRange(Check_026(formsList, currentFormLine));
            errorList.AddRange(Check_027(formsList, currentFormLine));

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

    //Проверка правильности заполнения дат периода
    private static List<CheckError> Check_001(Report rep)
    {
        List<CheckError> result = new();
        if (!DateOnly.TryParse(rep.StartPeriod_DB, out var stPer))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = string.Empty,
                Column = string.Empty,
                Value = rep.StartPeriod_DB,
                Message = "Некорректно введена дата начала периода."
            });
        }
        if (!DateOnly.TryParse(rep.EndPeriod_DB, out var endPer))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = string.Empty,
                Column = string.Empty,
                Value = rep.EndPeriod_DB,
                Message = "Некорректно введена дата окончания периода."
            });
        }
        if (DateOnly.TryParse(rep.StartPeriod_DB, out _)
            && DateOnly.TryParse(rep.EndPeriod_DB, out _)
            && stPer > endPer)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = string.Empty,
                Column = string.Empty,
                Value = $"{rep.StartPeriod_DB} - {rep.EndPeriod_DB}",
                Message = "Дата начала периода превышает дату окончания периода."
            });
        }
        return result;
    }

    #endregion

    #region Check002

    //Проверка даты окончания периода
    private static List<CheckError> Check_002(List<Form15> forms15, Report rep)
    {
        var forms = forms15.Cast<Form1>().ToList();
        return CheckRepPeriod(forms, rep);
    }

    #endregion

    #region Check003

    //Наличие строк дубликатов (графы 2-24)
    private static List<CheckError> Check_003(List<Form15> forms)
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
                                  && comparator.Compare(formToCompare.PassportNumber_DB, currentForm.PassportNumber_DB) == 0
                                  && comparator.Compare(formToCompare.Type_DB, currentForm.Type_DB) == 0
                                  && comparator.Compare(formToCompare.Radionuclids_DB, currentForm.Radionuclids_DB) == 0
                                  && comparator.Compare(formToCompare.FactoryNumber_DB, currentForm.FactoryNumber_DB) == 0
                                  && formToCompare.Quantity_DB == currentForm.Quantity_DB
                                  && comparator.Compare(formToCompare.Activity_DB, currentForm.Activity_DB) == 0
                                  && comparator.Compare(formToCompare.CreationDate_DB, currentForm.CreationDate_DB) == 0
                                  && comparator.Compare(formToCompare.StatusRAO_DB, currentForm.StatusRAO_DB) == 0
                                  && formToCompare.DocumentVid_DB == currentForm.DocumentVid_DB
                                  && comparator.Compare(formToCompare.DocumentNumber_DB, currentForm.DocumentNumber_DB) == 0
                                  && comparator.Compare(formToCompare.DocumentDate_DB, currentForm.DocumentDate_DB) == 0
                                  && comparator.Compare(formToCompare.ProviderOrRecieverOKPO_DB, currentForm.ProviderOrRecieverOKPO_DB) == 0
                                  && comparator.Compare(formToCompare.TransporterOKPO_DB, currentForm.TransporterOKPO_DB) == 0
                                  && comparator.Compare(formToCompare.PackName_DB, currentForm.PackName_DB) == 0
                                  && comparator.Compare(formToCompare.PackType_DB, currentForm.PackType_DB) == 0
                                  && comparator.Compare(formToCompare.PackNumber_DB, currentForm.PackNumber_DB) == 0
                                  && comparator.Compare(formToCompare.StoragePlaceName_DB, currentForm.StoragePlaceName_DB) == 0
                                  && comparator.Compare(formToCompare.StoragePlaceCode_DB, currentForm.StoragePlaceCode_DB) == 0
                                  && comparator.Compare(formToCompare.RefineOrSortRAOCode_DB, currentForm.RefineOrSortRAOCode_DB) == 0
                                  && comparator.Compare(formToCompare.Subsidy_DB, currentForm.Subsidy_DB) == 0
                                  && comparator.Compare(formToCompare.FcpNumber_DB, currentForm.FcpNumber_DB) == 0;
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
                FormNum = "form_15",
                Row = duplicateLines,
                Column = "2 - 24",
                Value = "",
                Message = $"Данные граф 2-24 в строках {duplicateLines} продублированы. " +
                          $"Следует проверить правильность предоставления данных."
            });
        }
        return result;
    }

    #endregion

    #region Check004

    //Нумерация строк (графа 1)
    private static List<CheckError> Check_004(List<Form15> forms, int line)
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
                Message = "Номера строк должны располагаться по порядку, без пропусков или дублирования номеров."
            });
        }
        return result;
    }

    #endregion

    #region Check005

    //Допустимый код операции (графа 2)
    private static List<CheckError> Check_005(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB ?? string.Empty;
        var operationCodeValid = new []
        {
            "01","10","14","21","22","25","26","27","28","29","31","32","35","36","37",
            "38","39","41","43","44","45","49","51","52","57","59","63","64","71","72",
            "73","74","75","76","84","88","97","98","99"
        };
        var valid = operationCodeValid.Contains(operationCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = "Код операции не может быть использован в форме 1.5."
            });
        }
        return result;
    }

    #endregion

    #region Check005_10

    //TODO
    //Проверка на соответствие СНК
    private static List<CheckError> Check_005_10(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "10" };
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
                Message = "Сведения, представленные в инвентаризации, не соответствуют СНК."
            });
        }
        return result;
    }

    #endregion
    
    #region Check005_29

    //Проверка наличия примечания (графа 2)
    private static List<CheckError> Check_005_29(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "29","39","49","59","97","98","99" };
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        const byte graphNumber = 2;
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = "В примечании к ячейке \"Код операции\" необходимо дать пояснение об осуществлённой операции."
            });
        }
        return result;
    }

    #endregion

    #region Check005_21

    //TODO
    //Проверка наличия учётной единицы в организации
    private static List<CheckError> Check_005_21(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[]
        {
            "21","22","25","26","27","28","29","42","43","44","45","49","51","71","72","84","98"
        };
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
                Message = "Учетной единицы с такими параметрами нет в организации. " +
                          "Проверьте правильность указываемых сведений для ОЗРИ."
            });
        }
        return result;
    }

    #endregion

    #region Check005_37

    //TODO
    //Проверка наличия кода операции 27 ранее
    private static List<CheckError> Check_005_37(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "37" };
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
                Message = "В отчетах не найдена строка об осуществлении передачи учетной единицы. " +
                          "Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check005_41

    //TODO
    //Эта проверка по идее никогда не выдаст ошибку, т.к. проверке все равно, появился ли 41 код вручную или автоматически.
    private static List<CheckError> Check_005_41(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "41" };
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
                Message = ""
            });
        }
        return result;
    }

    #endregion

    #region Check005_51

    //TODO
    //Проверка наличия кода операции 52
    private static List<CheckError> Check_005_51(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "51" };
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
                Message = ""
            });
        }
        return result;
    }

    #endregion

    #region Check005_52

    //TODO
    //Проверка наличия кода операции 51
    private static List<CheckError> Check_005_52(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "52" };
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
                Message = "В отчетах не найдена строка об изъятии РАО из пункта хранения. " +
                          "Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check005_57

    //TODO
    //Проверка кода операции 45
    private static List<CheckError> Check_005_57(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "57" };
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
                Message = "В отчетах не найдена строка снятии учетной единицы для упаковки/переупаковки. " +
                          "Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check005_84

    //Статус РАО равен 7 (графа 11)
    private static List<CheckError> Check_005_84(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "84" };
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
                Message = "Для ОЗРИ, которые возвращаются в страну поставщика код статус РАО – 7."
            });
        }
        return result;
    }

    #endregion

    #region Check005_71

    //Справочная "ошибка" - т.е. даже не ошибка.
    private static List<CheckError> Check_005_71(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var applicableOperationCodes = new[] { "71", "72", "73", "74", "75", "76" };
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
                Message = "К отчету необходимо приложить скан-копию документа, характеризующего операцию."
            });
        }
        return result;
    }

    #endregion

    #region Check005_9

    //Для кодов операции оканчивающихся на 9 должно быть примечание.
    private static List<CheckError> Check_005_9(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        const byte graphNumber = 2;
        var applicableOperationCodes = new[] { "29", "39", "49", "59", "97", "98", "99" };
        if (!applicableOperationCodes.Contains(opCode)) return result;
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = opCode,
                Message = "В примечании к ячейке \"Код операции\" необходимо дать пояснение об осуществлённой операции."
            });
        }
        return result;
    }

    #endregion

    #region Check006

    //Дата операции входит в отчетный период (колонка 3)
    private static List<CheckError> Check_006(List<Form15> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] nonApplicableOperationCodes = { "01" };
        var documentDate = forms[line].DocumentDate_DB;
        if (nonApplicableOperationCodes.Contains(forms[line].OperationCode_DB)
            || !DateOnly.TryParse(rep.StartPeriod_DB, out var pStart)
            || !DateOnly.TryParse(rep.EndPeriod_DB, out var pEnd)
            || !DateOnly.TryParse(forms[line].OperationDate_DB, out var opDate))
        {
            return result;
        }
        var valid = opDate >= pStart && opDate <= pEnd;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(opDate),
                Message = "Дата операции не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion

    #region Check006

    //Дата документа входит в отчетный период (колонка 3)
    private static List<CheckError> Check_006_03(List<Form15> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "10" };
        var documentDate = forms[line].DocumentDate_DB;
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)
            || !DateOnly.TryParse(rep.StartPeriod_DB, out var pStart)
            || !DateOnly.TryParse(rep.EndPeriod_DB, out var pEnd)
            || !DateOnly.TryParse(forms[line].DocumentDate_DB, out var docDate))
        {
            return result;
        }
        var valid = docDate >= pStart && docDate <= pEnd;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(docDate),
                Message = "Дата документа не входит в отчётный период. Для операции инвентаризации, " +
                          "срок предоставления отчёта исчисляется с даты утверждения акта инвентаризации."
            });
        }
        return result;
    }

    #endregion

    #region Check007

    //Номер паспорта заполнен (графа 4)
    private static List<CheckError> Check_007(List<Form15> forms, int line)
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
                Message = "Заполните сведения о номере паспорта (сертификата) ЗРИ, который переведен в ОЗИИИ."
            });
        }
        return result;
    }

    #endregion

    #region Check008

    //Тип заполнен (графа 5)
    private static List<CheckError> Check_008(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var type = forms[line].Type_DB;
        var valid = !string.IsNullOrWhiteSpace(type);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Type_DB",
                Value = Convert.ToString(type),
                Message = "Заполните сведения о типе ЗРИ, который переведен в ОЗИИИ."
            });
        }
        return result;
    }

    #endregion

    #region Check009_a

    //Радионуклиды заполнены (графа 6)
    private static List<CheckError> Check_009_a(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var radionuclids = forms[line].Radionuclids_DB;
        var valid = !string.IsNullOrWhiteSpace(radionuclids);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = Convert.ToString(radionuclids),
                Message = "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check009_b

    //Все радионуклиды есть в справочнике (графа 6)
    private static List<CheckError> Check_009_b(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var rads = forms[line].Radionuclids_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(rads) || rads.Trim() == "-") return result;
        var radsSet = rads
            .ToLower()
            .Replace(" ", string.Empty)
            .Replace(',', ';')
            .Split(';')
            .ToHashSet();
        if (radsSet.Count == 1 && R.All(phEntry => phEntry["name"] != radsSet.First()))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = rads,
                Message = "Формат ввода данных не соответствует приказу. " +
                          "Наименование радионуклида указывается названием химического элемента на русском языке " +
                          "с указанием через дефис массового числа изотопа."
            });
            return result;
        }
        if (rads.Contains(',')
            || rads.Contains('+')
            || !radsSet
                .All(rad => R
                    .Any(phEntry => phEntry["name"] == rad)))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Radionuclids_DB",
                Value = rads,
                Message = "Формат ввода данных не соответствует приказу. " +
                          "Наименование радионуклида указывается названием химического элемента на русском языке " +
                          "с указанием через дефис массового числа изотопа, радионуклиды перечисляются через точку с запятой."
            });
        }
        return result;
    }

    #endregion

    #region Check010_a

    //У номеров в качестве разделителя использовать ";"
    private static List<CheckError> Check_010_a(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var numbers = forms[line].FactoryNumber_DB ?? string.Empty;
        if (numbers.Contains(',') || numbers.Contains('+'))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "FactoryNumber_DB",
                Value = Convert.ToString(numbers),
                Message = "Заполните сведения о заводском номере ЗРИ, который переведен в ОЗИИИ. Если номер отсутствует, " +
                          "в ячейке следует указать символ \"-\" без кавычек. Для упаковки однотипных ЗРИ, " +
                          "имеющей один паспорт (сертификат) заводские номера в списке разделяются точкой с запятой."
            });
        }
        return result;
    }

    #endregion

    #region Check010_b

    //Номер не пустая строка (колонка 7)
    private static List<CheckError> Check_010_b(List<Form15> forms, int line)
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
                Message = "Заполните сведения о заводском номере ЗРИ, который переведен в ОЗИИИ. Если номер отсутствует, " +
                          "в ячейке следует указать символ \"-\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check011

    //Количество заполнено (графа 8)
    private static List<CheckError> Check_011(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var factoryNum = forms[line].FactoryNumber_DB;
        var quantity = forms[line].Quantity_DB ?? 0;
        if (factoryNum == null) return result;
        if (quantity is 0)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Quantity_DB",
                Value = Convert.ToString(quantity),
                Message = "Заполните сведения о количестве ЗРИ, переведенных в ОЗИИИ."
            });
            return result;
        }
        var valid = !string.IsNullOrWhiteSpace(factoryNum)
                    && !factoryNum.Contains(',');
        if (quantity > 1 && !factoryNum.Contains(';')
            || quantity == 1 && factoryNum.Contains(';'))
        {
            valid = false;
        }
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Quantity_DB",
                Value = quantity.ToString(),
                Message = "Формат ввода данных не соответствует приказу. Номера ЗРИ должны быть разделены точкой с запятой."
            });
        }
        return result;
    }

    #endregion

    #region Check012

    //Проверка на корректные данные суммарной активности (графа 9)
    private static List<CheckError> Check_012(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var activity = ConvertStringToExponential(forms[line].Activity_DB);
        if (string.IsNullOrEmpty(activity) 
            || activity == "-"
            || !double.TryParse(activity,
                NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
                CultureInfo.CreateSpecificCulture("ru-RU"),
                out var activityReal))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = Convert.ToString(forms[line].Activity_DB),
                Message = "Заполните сведения о суммарной активности ЗРИ, переведенных в ОЗИИИ. " +
                          "Оценочные сведения приводятся в круглых скобках."
            });
        }
        else if (activityReal < 10)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = Convert.ToString(forms[line].Activity_DB),
                Message = "Суммарная активность должна быть более 10 Бк. Проверьте правильность введённых данных."
            });
        }
        else if (activityReal > 10e+20)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Activity_DB",
                Value = Convert.ToString(forms[line].Activity_DB),
                Message = "Указано слишком большое значение суммарной активности. Проверьте правильность введённых данных."
            });
        }
        return result;
    }
    #endregion

    #region Check013

    //Дата выпуска <= дате операции
    private static List<CheckError> Check_013(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationDate = forms[line].OperationDate_DB ?? string.Empty;
        var creationDate = forms[line].CreationDate_DB ?? string.Empty;
        if (!DateTime.TryParse(creationDate, out var creationDateReal) && creationDate != "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "CreationDate_DB",
                Value = Convert.ToString(creationDate),
                Message = "Дата выпуска заполнена некорректно."
            });
            return result;
        }
        if (!DateTime.TryParse(operationDate, out var operationDateReal))
        {
            return result;
        }
        var valid = creationDateReal <= operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "CreationDate_DB",
                Value = Convert.ToString(creationDate),
                Message = "Дата выпуска не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check014

    //Проверка соответствия кода операции и статуса РАО (графа 11)
    private static List<CheckError> Check_014(List<Form15> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        var status = forms[line].StatusRAO_DB ?? string.Empty;
        var jurOkpo = forms10[0].Okpo_DB ?? string.Empty;
        var obOkpo = forms10[1].Okpo_DB ?? string.Empty;

        switch (opCode)
        {
            case "11" or "12" or "13" or "14" or "41":
            {
                if (status != obOkpo && status != jurOkpo)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "РАО, образовавшиеся после 15.07.2011, находятся в собственности организации, " +
                                  "в результате деятельности которой они образовались."
                    });
                }
                break;
            }
            case "26" or "28" or "63":
            {
                if (status != obOkpo && status != jurOkpo)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "Операция, соответствующая выбранному коду, может использоваться только для собственных РАО."
                    });
                }
                break;
            }
            case "38" or "64":
            {
                if (status != obOkpo && status != jurOkpo)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "При операциях, связанных с получением права собственности, " +
                                  "в графе статус РАО необходимо отразить код ОКПО отчитывающейся организации."
                    });
                }
                break;
            }
            case "42" or "43" or "73" or "97" or "98":
            {
                if (status != obOkpo && status != jurOkpo)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "Проверьте правильность статуса РАО."
                    });
                }
                break;
            }
            case "22" or "32":
            {
                if (status != "1" && !OkpoRegex.IsMatch(status))
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "Проверьте правильность статуса РАО."
                    });
                }
                break;
            }
            case "16":
            {
                if (status != "2" && !OkpoRegex.IsMatch(status))
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "Проверьте правильность статуса РАО."
                    });
                }
                break;
            }
            case "76":
            {
                if (status is not ("6" or "9"))
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "Проверьте правильность статуса РАО."
                    });
                }
                break;
            }
            default:
            {
                if (status is not ("1" or "2" or "3" or "4" or "6" or "9") && !OkpoRegex.IsMatch(status))
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "StatusRAO_DB",
                        Value = Convert.ToString(status),
                        Message = "Заполнение графы 5 не соответствует требованиям приказа Госкорпорации \"Росатом\" " +
                                  "от 07.12.2020 №1/13-НПА."
                    });
                }
                break;
            }
        }
        return result;
    }

    #endregion

    #region Check015

    //Проверка кода вида документа 
    private static List<CheckError> Check_015(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 12;
        var documentVid = forms[line].DocumentVid_DB;
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        var valid = documentVid is >= 1 and <= 15 or 19;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = "Неверный код вида документа, сопровождающего операцию."
            });
            return result;
        }
        if (documentVid == 19)
        {
            valid = CheckNotePresence(notes, line, graphNumber);
            if (!valid)
            {
                result.Add(new CheckError
                {
                    FormNum = "form_15",
                    Row = (line + 1).ToString(),
                    Column = "DocumentVid_DB",
                    Value = Convert.ToString(documentVid),
                    Message = "К коду вида документа 19 необходимо примечание."
                });
            }
        } 
        else if (opCode == "41" && documentVid != 1)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = "При коде операции 31, вид документа должен быть равен 1."
            });
        }
        return result;
    }

    #endregion

    #region Check016

    //Номер документа заполнен
    private static List<CheckError> Check_016(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var documentNumber = forms[line].DocumentNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(documentNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentNumber_DB",
                Value = Convert.ToString(documentNumber),
                Message = "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check017_a

    //Дата документа корректно заполнена (графа 17)
    private static List<CheckError> Check_017_a(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var docDate = forms[line].DocumentDate_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(docDate) || docDate is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(docDate),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        else if (!DateOnly.TryParse(docDate, out _))
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(docDate),
                Message = "Формат ввода данных не соответствует приказу. Некорректно заполнена дата документа."
            });
        }
        return result;
    }

    #endregion

    #region Check017_b

    //Дата документа <= дате операции для всех кодов операции, кроме 10 и 41
    private static List<CheckError> Check_017_b(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB ?? string.Empty;
        var documentDate = forms[line].DocumentDate_DB ?? string.Empty;
        if (operationCode is "10" or "41"
            || !DateOnly.TryParse(documentDate, out var documentDateReal)
            || !DateOnly.TryParse(operationDate, out var operationDateReal))
        {
            return result;
        }
        var valid = documentDateReal <= operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(documentDate),
                Message = "Дата документа не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check017_c

    //Дата документа = дате операции, если код операции 41
    private static List<CheckError> Check_017_c(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var documentDate = forms[line].DocumentDate_DB;
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        if (operationCode is not "41"
            || !DateOnly.TryParse(documentDate, out var documentDateReal)
            || !DateOnly.TryParse(operationDate, out var operationDateReal))
        {
            return result;
        }
        var valid = documentDateReal == operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(documentDate),
                Message = "Дата документа должна соответствовать дате операции."
            });
        }
        return result;
    }

    #endregion

    #region Check017_d

    //При коде операции 10, дата окончания ОП не позднее даты документа + 10 дней
    private static List<CheckError> Check_017_d(List<Form15> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        var documentDate = forms[line].DocumentDate_DB ?? string.Empty;
        var endPeriod = rep.EndPeriod_DB ?? string.Empty;

        if (opCode is not "10"
            || !DateOnly.TryParse(endPeriod, out var pEnd)
            || !DateOnly.TryParse(documentDate, out var docDate)) return result;

        var valid = WorkdaysBetweenDates(docDate, pEnd) <= 10;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата документа не входит в отчетный период. Для операции инвентаризации, " +
                          "срок предоставления отчета исчисляется с даты утверждения акта инвентаризации."
            });
        }
        return result;
    }

    #endregion

    #region Check018_01

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (графа 18)
    private static List<CheckError> Check_018_01(List<Form15> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "01", "10", "14", "41", "43", "44", "45", "49", "51", "52", "57", 
            "59", "71", "72", "73", "74", "75", "76", "97", "98", "99"
        };
        var operationCode = forms[line].OperationCode_DB ?? string.Empty;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB ?? string.Empty;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        
        var valid = providerOrRecieverOkpo == repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Для выбранного кода операции указывается код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check018_84

    //Код ОКПО поставщика/получателя из ОКСМ (не Россия), с примечанием (графа 18)
    private static List<CheckError> Check_018_84(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 15;
        string[] applicableOperationCodes = { "84", "88" };
        var operationCode = forms[line].OperationCode_DB ?? string.Empty;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB ?? string.Empty;
        
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == providerOrRecieverOkpo)
                    && !providerOrRecieverOkpo.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "При операциях, связанных с перемещением ОЗИИИ через государственную границу Российской Федерации, " +
                          "необходимо указывать краткое наименование государства в соответствии с ОКСМ."
            });
        }
        valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Необходимо добавить примечание."
            });
        }
        return result;
    }

    #endregion

    #region Check018_21

    //Код ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации + 8/14 цифр (колонка 18)
    private static List<CheckError> Check_018_21(List<Form15> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "21", "25", "26", "27", "28", "29", "31", "35", "36", "37", "38", "39" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Значение может состоять только из 8 или 14 символов"
            });
        }
        valid = providerOrRecieverOkpo != repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Для выбранного кода операции указывается код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check018_22

    //Код ОКПО поставщика/получателя состоит из 8/14 чисел или "минобороны" (графа 18)
    private static List<CheckError> Check_018_22(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        if (operationCode is not ("22" or "32")) return result;
        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo)
                    || providerOrRecieverOkpo.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = Convert.ToString(providerOrRecieverOkpo),
                Message = "Формат ввода данных не соответствует приказу. Следует указать код ОКПО контрагента, " +
                          "либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check019_01

    //Код ОКПО перевозчика равен "-" (графа 16)
    private static List<CheckError> Check_019_01(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = 
        {
            "01", "10", "14", "18", "41", "43", "44", "45", "48", "49", "51", 
            "52", "57", "59", "71", "72", "73", "74", "75", "76", "97", "98"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOkpo = forms[line].TransporterOKPO_DB ?? string.Empty;
        var valid = transporterOkpo is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOkpo),
                Message = "При выбранном коде операции транспортирование не производится."
            });
        }
        return result;
    }

    #endregion

    #region Check019_21 (графа 16)

    //При определенных кодах операции, код ОКПО перевозчика равен 8/14 цифр
    private static List<CheckError> Check_019_21(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21","25","26","27","28","29","31","35","36","37","38","39","84","85","86","88"
        };
        var operationCode = forms[line].OperationCode_DB ?? string.Empty;
        var transporterOkpo = forms[line].TransporterOKPO_DB ?? string.Empty;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(transporterOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOkpo),
                Message = "Необходимо указать код ОКПО организации перевозчика."
            });
        }
        return result;
    }

    #endregion

    #region Check019_22 (графа 16)

    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны"
    private static List<CheckError> Check_019_22(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "22", "32" };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOkpo = forms[line].TransporterOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(transporterOkpo)
                    || transporterOkpo.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = Convert.ToString(transporterOkpo),
                Message = "Необходимо указать код ОКПО организации перевозчика, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check020

    //Проверка Наименования прибора (графа 17)
    private static List<CheckError> Check_020(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var packName = forms[line].PackName_DB ?? string.Empty;
        var packType = forms[line].PackType_DB ?? string.Empty;
        var packNum = forms[line].PackNumber_DB ?? string.Empty;
        if (packName.ToLower() is "без упаковки" && packType is not "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackType_DB",
                Value = Convert.ToString(packType),
                Message = "В случае, если упаковка отсутствует, в графе \"Тип прибора\" должен быть указан \"-\" без кавычек."
            });
        }
        if (packName.ToLower() is "без упаковки" && packNum is not "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackNumber_DB",
                Value = Convert.ToString(packNum),
                Message = "В случае, если упаковка отсутствует, " +
                          "в графе \"Заводской номер прибора\" должен быть указан \"-\" без кавычек."
            });
        }
        var valid = !string.IsNullOrWhiteSpace(packName) && packName is not "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackName_DB",
                Value = Convert.ToString(packName),
                Message = "Заполните сведения об упаковке РАО. Если РАО размещены без упаковки, " +
                          "то в графе 17 указывается \"без упаковки\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check021

    //Тип прибора заполнен (графа 18)
    private static List<CheckError> Check_021(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var packType = forms[line].PackType_DB ?? string.Empty;
        var valid = !string.IsNullOrWhiteSpace(packType);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackType_DB",
                Value = Convert.ToString(packType),
                Message = "Заполните сведения в графе \"Тип прибора\". " +
                          "В случае, если тип отсутствует, укажите символ \"=\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check022

    //Номер прибора заполнен (графа 19)
    private static List<CheckError> Check_022(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var packNum = forms[line].PackNumber_DB ?? string.Empty;
        var valid = !string.IsNullOrWhiteSpace(packNum);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "PackNumber_DB",
                Value = Convert.ToString(packNum),
                Message = "Заполните сведения о заводском номере упаковки. " +
                          "Если заводской номер отсутствует необходимо привести в круглых скобках номер, присвоенный в организации."
            });
        }
        return result;
    }

    #endregion

    #region Check023

    //Наименование пункта хранения заполнено (графа 20)
    private static List<CheckError> Check_023(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var storagePlaceName = forms[line].StoragePlaceName_DB ?? string.Empty;
        var valid = !string.IsNullOrWhiteSpace(storagePlaceName);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceName_DB",
                Value = Convert.ToString(storagePlaceName),
                Message = "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check024

    //Код пункта хранения заполнен (графа 21)
    private static List<CheckError> Check_024(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var storagePlaceCode = forms[line].StoragePlaceCode_DB;
        var valid = !string.IsNullOrWhiteSpace(storagePlaceCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "StoragePlaceCode_DB",
                Value = Convert.ToString(storagePlaceCode),
                Message = "Графа должна быть заполнена."
            });
        }
        return result;
    }

    #endregion

    #region Check025

    //
    private static List<CheckError> Check_025(List<Form15> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var sortCode = forms[line].RefineOrSortRAOCode_DB ?? string.Empty;
        var valid = !string.IsNullOrWhiteSpace(sortCode);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "RefineOrSortRAOCode_DB",
                Value = Convert.ToString(sortCode),
                Message = "Графа должна быть заполнена."
            });
        }
        else
        {
            var operationCode = forms[line].OperationCode_DB ?? string.Empty;

            switch (operationCode)
            {
                case "01" or "10" or "14" or "21" or "22" or "25" or "26" or "27" or "28" or "29" or "31" 
                    or "32" or "35" or "36" or "37" or "38" or "39" or "43" or "51" or "52" or "63" or "64" 
                    or "71" or "72" or "73" or "74" or "75" or "76" or "84" or "88" or "97" or "98" or "99":
                {
                    valid = sortCode == "-";
                    if (!valid)
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_15",
                            Row = (line + 1).ToString(),
                            Column = "RefineOrSortRAOCode_DB",
                            Value = Convert.ToString(sortCode),
                            Message = "Выбранный код операции не соответствует переработке/сортировке РАО. " +
                                      "Проверьте правильность заполнения граф 2 и 22."
                        });
                    }
                    break;
                }
                case "44":
                {
                    if (sortCode is not ("41" or "42" or "43" or "49" or "54" or "71" or "72" or "73" or "74" or "79" or "99"))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_15",
                            Row = (line + 1).ToString(),
                            Column = "RefineOrSortRAOCode_DB",
                            Value = Convert.ToString(sortCode),
                            Message = "Несуществующий код переработки/сортировки."
                        });
                    }

                    break;
                }
                case "45" or "57":
                {
                    if (sortCode is not ("-" or "74"))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_15",
                            Row = (line + 1).ToString(),
                            Column = "RefineOrSortRAOCode_DB",
                            Value = Convert.ToString(sortCode),
                            Message = "Коду операции упаковка/переупаковка не соответствует код переработки/сортировки."
                        });
                    }

                    break;
                }
                case "49" or "59":
                {
                    if (sortCode is not ("-" or "52" or "72" or "74"))
                    {
                        result.Add(new CheckError
                        {
                            FormNum = "form_15",
                            Row = (line + 1).ToString(),
                            Column = "RefineOrSortRAOCode_DB",
                            Value = Convert.ToString(sortCode),
                            Message = "Коду операции сортировка соответствуют коды сортировки 52, 72, 74."
                        });
                    }

                    break;
                }
            }
            if (sortCode is "49" or "79" or "99")
            {
                const byte graphNumber = 22;
                valid = CheckNotePresence(notes, line, graphNumber);
                if (!valid)
                {
                    result.Add(new CheckError
                    {
                        FormNum = "form_15",
                        Row = (line + 1).ToString(),
                        Column = "RefineOrSortRAOCode_DB",
                        Value = Convert.ToString(sortCode),
                        Message = "К выбранному коду переработки/сортировки нужно привести примечание."
                    });
                }
            }
        }
        return result;
    }

    #endregion

    #region Check026

    //Субсидия от 0 до 100, либо "-", либо "" (графа 23)
    private static List<CheckError> Check_026(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var subsidy = ConvertStringToExponential(forms[line].Subsidy_DB);
        var valid = subsidy is "-" or "" 
                    || (double.TryParse(subsidy,
                            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
                            CultureInfo.CreateSpecificCulture("ru-RU"),
                            out var subsidyNum)
                        && subsidyNum is >= 0 and <= 100);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "Subsidy_DB",
                Value = Convert.ToString(subsidy),
                Message = "Графа заполнена некорректно. Допустимы значения от 0 до 100, символ \"-\" без кавычек или пустая строка."
            });
        }
        return result;
    }

    #endregion

    #region Check027

    //Номер ФЦП заполнен (графа 24)
    private static List<CheckError> Check_027(List<Form15> forms, int line)
    {
        List<CheckError> result = new();
        var fcpNum = ConvertStringToExponential(forms[line].FcpNumber_DB);
        var valid = fcpNum is "-" or ""
                    || (double.TryParse(fcpNum,
                            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
                            CultureInfo.CreateSpecificCulture("ru-RU"),
                            out _));
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_15",
                Row = (line + 1).ToString(),
                Column = "FcpNumber_DB",
                Value = Convert.ToString(fcpNum),
                Message = "Графа заполнена некорректно. Допустимо цифровое значение, символ \"-\" без кавычек или пустая строка."
            });
        }
        return result;
    }

    #endregion

    #endregion

    #region GraphList

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
}
