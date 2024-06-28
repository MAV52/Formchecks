using Models.CheckForm;
using Models.Collections;
using Models.Forms;
using Models.Forms.Form1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Reflection;

namespace Client_App.Commands.SyncCommands.CheckForm;

public class CheckF12 : CheckBase
{
    #region CheckTotal

    public static List<CheckError> Check_Total(Reports reps, Report rep)
    {
        var currentFormLine = 0;
        List<CheckError> errorList = new();
        LoadDictionaries();
        var formsList = rep.Rows12.ToList<Form12>();
        errorList.AddRange(Check_002(rep));
        errorList.AddRange(Check_003(formsList, rep));
        errorList.AddRange(Check_004(formsList));
        foreach (var key in rep.Rows12)
        {
            var form = (Form12)key;
            var notes = rep.Notes.ToList<Note>();
            var forms10 = reps.Master_DB.Rows10.ToList<Form10>();
            errorList.AddRange(Check_001(formsList, currentFormLine));
            errorList.AddRange(Check_005(formsList, currentFormLine));
            errorList.AddRange(Check_006(formsList, currentFormLine));
            errorList.AddRange(Check_007(formsList, currentFormLine));
            errorList.AddRange(Check_008(formsList, currentFormLine));
            errorList.AddRange(Check_009(formsList, notes, currentFormLine));
            errorList.AddRange(Check_010(formsList, currentFormLine));
            errorList.AddRange(Check_011(formsList, currentFormLine));
            errorList.AddRange(Check_012(formsList, currentFormLine));
            errorList.AddRange(Check_013(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_014(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_015(formsList, currentFormLine));
            errorList.AddRange(Check_016(formsList, currentFormLine));
            errorList.AddRange(Check_017(formsList, currentFormLine));
            errorList.AddRange(Check_018(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_019(formsList, currentFormLine));
            errorList.AddRange(Check_020(formsList, currentFormLine));
            errorList.AddRange(Check_021(formsList, currentFormLine));
            errorList.AddRange(Check_022(formsList, rep, currentFormLine));
            errorList.AddRange(Check_023(formsList, rep, currentFormLine));
            errorList.AddRange(Check_024(formsList, currentFormLine));
            errorList.AddRange(Check_025(formsList, currentFormLine));
            errorList.AddRange(Check_026(formsList, currentFormLine));
            errorList.AddRange(Check_027(formsList, currentFormLine));
            errorList.AddRange(Check_028(formsList, currentFormLine));
            errorList.AddRange(Check_029(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_030(formsList, currentFormLine));
            errorList.AddRange(Check_031(formsList, notes, currentFormLine));
            errorList.AddRange(Check_032(formsList, currentFormLine));
            errorList.AddRange(Check_033(formsList, currentFormLine));
            errorList.AddRange(Check_034(formsList, currentFormLine));
            errorList.AddRange(Check_035(formsList, currentFormLine));
            errorList.AddRange(Check_036(formsList, notes, currentFormLine));
            errorList.AddRange(Check_037(formsList, notes, currentFormLine));
            errorList.AddRange(Check_038(formsList, notes, currentFormLine));
            errorList.AddRange(Check_039(formsList, currentFormLine));
            errorList.AddRange(Check_040(formsList, currentFormLine));
            errorList.AddRange(Check_041(formsList, notes, currentFormLine));
            errorList.AddRange(Check_042(formsList, currentFormLine));
            errorList.AddRange(Check_043(formsList, currentFormLine));
            errorList.AddRange(Check_044(formsList, currentFormLine));
            errorList.AddRange(Check_045(formsList, currentFormLine));
            errorList.AddRange(Check_046(formsList, currentFormLine));
            errorList.AddRange(Check_047(formsList, currentFormLine));
            errorList.AddRange(Check_048(formsList, currentFormLine));
            errorList.AddRange(Check_049(formsList, rep, currentFormLine));
            errorList.AddRange(Check_050(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_051(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_052(formsList, currentFormLine));
            errorList.AddRange(Check_053(formsList, forms10, currentFormLine));
            errorList.AddRange(Check_054(formsList, currentFormLine));
            errorList.AddRange(Check_055(formsList, notes, currentFormLine));
            errorList.AddRange(Check_056(formsList, currentFormLine));
            errorList.AddRange(Check_057(formsList, currentFormLine));
            errorList.AddRange(Check_058(formsList, currentFormLine));
            errorList.AddRange(Check_059(formsList, currentFormLine));
            errorList.AddRange(Check_060(formsList, currentFormLine));
            errorList.AddRange(Check_061(formsList, currentFormLine));

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

    //Заглушка
    private static List<CheckError> Check_001(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        return result;
    }

    #endregion

    #region Check002

    //Проверка правильности заполнения дат периода
    private static List<CheckError> Check_002(Report rep)
    {
        List<CheckError> result = new();
        if (!DateOnly.TryParse(rep.StartPeriod_DB, out var stPer))
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
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
                FormNum = "form_12",
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
                FormNum = "form_12",
                Row = string.Empty,
                Column = string.Empty,
                Value = $"{rep.StartPeriod_DB} - {rep.EndPeriod_DB}",
                Message = "Дата начала периода превышает дату окончания периода."
            });
        }
        return result;
    }

    #endregion

    #region Check003

    //Проверка даты окончания периода
    private static List<CheckError> Check_003(List<Form12> forms12, Report rep)
    {
        var forms = forms12.Cast<Form1>().ToList();
        return CheckRepPeriod(forms, rep);
    }

    #endregion

    #region Check004

    //Наличие строк дубликатов (графы 2-16)
    private static List<CheckError> Check_004(List<Form12> forms)
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
                                  && comparator.Compare(formToCompare.NameIOU_DB, currentForm.NameIOU_DB) == 0
                                  && comparator.Compare(formToCompare.FactoryNumber_DB, currentForm.FactoryNumber_DB) == 0
                                  && comparator.Compare(formToCompare.Mass_DB, currentForm.Mass_DB) == 0
                                  && comparator.Compare(formToCompare.CreatorOKPO_DB, currentForm.CreatorOKPO_DB) == 0
                                  && comparator.Compare(formToCompare.CreationDate_DB, currentForm.CreationDate_DB) == 0
                                  && comparator.Compare(formToCompare.SignedServicePeriod_DB, currentForm.SignedServicePeriod_DB) == 0
                                  && formToCompare.PropertyCode_DB == currentForm.PropertyCode_DB
                                  && comparator.Compare(formToCompare.Owner_DB, currentForm.Owner_DB) == 0
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
                FormNum = "form_12",
                Row = duplicateLines,
                Column = "2 - 16",
                Value = "",
                Message = $"Данные граф 2-16 в строках {duplicateLines} продублированы. Следует проверить правильность предоставления данных."
            });
        }
        return result;
    }

    #endregion

    #region Check005

    //Порядок строк (графа 1)
    private static List<CheckError> Check_005(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (forms[line].NumberInOrder_DB != line + 1)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Id",
                Value = forms[line].Id.ToString(),
                Message = "Номера строк должны располагаться по порядку, без пропусков или дублирования номеров."
            });
        }
        return result;
    }

    #endregion

    #region Check006

    //Код операции из списка (графа 2)
    private static List<CheckError> Check_006(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var operationCodeDbValid = new []
        {
            "10","11","12","17","18","21","22","25","27","28","29","31","32","35","37","38","39",
            "41","42","46","53","54","58","61","62","63","64","66","67","68","71","72","73","74",
            "75","81","82","83","84","85","86","87","88","97","98","99"
        };
        var operationCode = forms[line].OperationCode_DB ?? "";
        if (!operationCodeDbValid.Contains(operationCode) && operationCode != "43")
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = operationCode,
                Message = "Этот код операции не может быть использован в форме 1.2."
            });
        }
        return result;
    }

    #endregion

    #region Check007

    //Для кода операции 43 отдельное сообщение об ошибке (графа 3)
    private static List<CheckError> Check_007(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        if (opCode is "43")
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "Активность изделий из обедненного урана не может снизится за время эксплуатации " +
                          "до значений ниже уровня отнесения к объектам учета."
            });
        }
        return result;
    }

    #endregion

    #region Check008

    //TODO
    //СНК
    private static List<CheckError> Check_008(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        return result;
    }

    #endregion

    #region Check009

    //Наличие примечания (графа 2)
    private static List<CheckError> Check_009(List<Form12> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 2;
        string[] applicableOperationCodes = { "29", "39", "97", "98", "99" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "Необходимо дать пояснение об осуществленной операции."
            });
        }
        return result;
    }

    #endregion

    #region Check010

    //TODO
    //Наличие учётной единицы в СНК
    private static List<CheckError> Check_010(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21", "22", "25", "27", "28", "29", "41", "42", "46", "53", "54", "61", "62", 
            "65", "66", "67", "68", "71", "72", "81", "82", "83", "84", "88", "98"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        // ...
        // ...
        // ...
        return result;
    }

    #endregion

    #region Check011

    //TODO
    //Наличие записи с кодом операции 27
    private static List<CheckError> Check_011(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "37" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        //..
        //..
        //..
        return result;
    }

    #endregion

    #region Check012

    //TODO
    //Наличие отчёта по форме 1.6
    private static List<CheckError> Check_012(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "41" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        //..
        //..
        //..
        return result;
    }

    #endregion

    #region Check013

    //ОКПО отчитывающейся организации в ОКПО поставщика или получателя (графа 16)
    private static List<CheckError> Check_013(List<Form12> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "53" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        var valid = providerOrRecieverOkpo == repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "В графе 16 необходимо указать код ОКПО отчитывающейся организации. " +
                          "В случае, если зарядку/разрядку осуществляла подрядная организация, " +
                          "следует использовать код операции 54."
            });
        }
        return result;
    }

    #endregion

    #region Check014

    //Не ОКПО отчитывающейся организации в ОКПО поставщика или получателя(графа 16)
    private static List<CheckError> Check_014(List<Form12> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "54" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        var valid = providerOrRecieverOkpo != repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "В графе 16 необходимо указать ОКПО подрядной организации. В случае, если зарядка/разрядка осуществлялась силами отчитывающейся организации, следует использовать код операции 53."
            });
        }
        return result;
    }

    #endregion

    #region Check015

    //TODO
    //Наличие кода операции 46
    private static List<CheckError> Check_015(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "58" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "В отчетах не найдены сведения о снятии с учета учетной единицы для разукомплектования. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check016

    //TODO
    //Наличие кода операции 61
    private static List<CheckError> Check_016(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "62" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "В отчетах не найдены сведения о вывозе учетной единицы. Проверьте правильность выбранного кода операции."
            });
        }
        return result;
    }

    #endregion

    #region Check017

    //TODO
    //Наличие отчёта по форме 1.3/1.4
    private static List<CheckError> Check_017(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (DB_Ignore) return result;
        string[] applicableOperationCodes = { "65" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var valid = false;
        //..
        //..
        //..
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "-",
                Value = "-",
                Message = "Заполните форму 1.3."
            });
        }
        return result;
    }

    #endregion

    #region Check018

    //Код ОКПО юридического лица или обособленного подразделения в код ОКПО правообладателя (графа 12)
    private static List<CheckError> Check_018(List<Form12> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "11", "12", "28", "38", "41", "63", "64", "73", "81", "85" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var owner = forms[line].Owner_DB;
        var okpoRepJur = forms10[0].Okpo_DB ?? "";
        var okpoRepTerPodr = forms10[1].Okpo_DB ?? "";
        var valid = !string.IsNullOrWhiteSpace(owner) 
                    && (owner == okpoRepTerPodr || owner == okpoRepJur);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Уточните правообладателя ИОУ."
            });
        }
        return result;
    }

    #endregion

    #region Check019

    //Код ОКПО изготовителя 8 или 14 цифр (графа 8)
    private static List<CheckError> Check_019(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "81", "88" };
        var operationCode = forms[line].OperationCode_DB;
        var creatorOkpo = forms[line].CreatorOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(creatorOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationCode_DB",
                Value = forms[line].OperationCode_DB,
                Message = "Код используется для предоставления сведений о ИОУ, произведенных в Российской Федерации."
            });
        }
        return result;
    }

    #endregion

    #region Check020

    //Код ОКПО изготовителя из ОКСМ (графа 8)
    private static List<CheckError> Check_020(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "83", "84", "85", "86" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        if (OKSM.All(oksmEntry => oksmEntry["shortname"] != forms[line].CreatorOKPO_DB)
            || forms[line].CreatorOKPO_DB.ToLower() is "россия")
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = forms[line].CreatorOKPO_DB,
                Message = "Код используется для предоставления сведений о ИОУ, произведенных за пределами Российской Федерации. " +
                          "Для импортированных ИОУ необходимо указать краткое наименование государства в соответствии с ОКСМ"
            });
        }
        return result;
    }

    #endregion

    #region Check021

    //Дата операции корректно заполнена (графа 3)
    private static List<CheckError> Check_021(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var opDate = forms[line].OperationDate_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(opDate) || opDate is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(opDate),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        else if (!DateOnly.TryParse(opDate, out _))
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = Convert.ToString(opDate),
                Message = "Формат ввода данных не соответствует приказу. Некорректно заполнена дата операции."
            });
        }
        return result;
    }

    #endregion

    #region Check022

    //Дата операции попадает в отчётный период (графа 3)
    private static List<CheckError> Check_022(List<Form12> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB;
        if (opCode is "10"
            || !DateOnly.TryParse(rep.StartPeriod_DB, out var pStart)
            || !DateOnly.TryParse(rep.EndPeriod_DB, out var pEnd)
            || !DateOnly.TryParse(forms[line].OperationDate_DB!, out var opDate))
        {
            return result;
        }

        var valid = opDate >= pStart && opDate <= pEnd;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "OperationDate_DB",
                Value = forms[line].OperationDate_DB,
                Message = "Дата операции не входит в отчетный период."
            });
        }
        return result;
    }

    #endregion

    #region Check023

    //При коде операции 10, дата документа должна попадать в отчетный период (графа 3)
    private static List<CheckError> Check_023(List<Form12> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var operationCode = forms[line].OperationCode_DB;
        if (operationCode is not "10"
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
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(forms[line].DocumentDate_DB),
                Message = "Дата документа не входит в отчетный период. Для операции инвентаризации срок предоставления отчета " +
                          "исчисляется с даты утверждения акта инвентаризации."
            });
        }
        return result;
    }

    #endregion

    #region Check024

    //Номер паспорта заполнен (графа 4)
    private static List<CheckError> Check_024(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var valid = !string.IsNullOrWhiteSpace(forms[line].PassportNumber_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PassportNumber_DB",
                Value = forms[line].PassportNumber_DB,
                Message = "Графа не может быть пустой. Укажите номера паспорта. " +
                          "В случае, если номер отсутствует укажите символ \"-\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check025

    //Наименование заполнено (графа 5)
    private static List<CheckError> Check_025(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var valid = !string.IsNullOrWhiteSpace(forms[line].NameIOU_DB);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "NameIOU_DB",
                Value = forms[line].NameIOU_DB,
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check026

    //Номер ИОУ заполнен, не содержит запятых (графа 6)
    private static List<CheckError> Check_026(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var factoryNum = forms[line].FactoryNumber_DB ?? "";
        var valid = !string.IsNullOrWhiteSpace(factoryNum) && !factoryNum.Contains(',');
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "FactoryNumber_DB",
                Value = forms[line].FactoryNumber_DB,
                Message = "Заполните сведения о заводском номере ИОУ. " +
                          "Если номер отсутствует, в ячейке следует указать символ \"-\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check027

    //Масса заполнена, больше нуля (графа 7)
    private static List<CheckError> Check_027(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var mass = ConvertStringToExponential(forms[line].Mass_DB);
        double.TryParse(mass,
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
            CultureInfo.CreateSpecificCulture("ru-RU"),
            out var massTrue);
        var valid = !string.IsNullOrWhiteSpace(forms[line].Mass_DB) && massTrue > 0;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = mass,
                Message = "Графа не может быть пустой. При отсутствии сведений - в круглых скобках указывается масса, " +
                          "установленная путем измерений."
            });
        }
        return result;
    }

    #endregion

    #region Check028

    //Масса меньше 10е05 (графа 7)
    private static List<CheckError> Check_028(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (string.IsNullOrWhiteSpace(forms[line].Mass_DB)) return result;
        var mass = ConvertStringToExponential(forms[line].Mass_DB);
        var valid = double.TryParse(mass,
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign,
            CultureInfo.CreateSpecificCulture("ru-RU"),
            out var massTrue);
        if (!valid || massTrue > 10e05)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Mass_DB",
                Value = mass,
                Message = "Проверьте правильность предоставления сведений о массе ИОУ."
            });
        }
        return result;
    }

    #endregion

    #region Check029

    //Код ОКПО изготовителя равен коду ОКПО организации (графа 8)
    private static List<CheckError> Check_029(List<Form12> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "11" };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var creatorOkpo = forms[line].CreatorOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        var valid = creatorOkpo == repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = creatorOkpo,
                Message = "Проверьте код ОКПО организации-изготовителя."
            });
        }
        return result;
    }

    #endregion

    #region Check030

    //Если в коде ОКПО изготовителя цифры, то это 8 или 14 цифр (графа 8)
    private static List<CheckError> Check_030(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var okpoRegexApplicable = new Regex(@"^\d+[_]?\d*$");
        var creatorOkpo = forms[line].CreatorOKPO_DB;
        if (!okpoRegexApplicable.IsMatch(creatorOkpo)) return result;
        var valid = !string.IsNullOrEmpty(creatorOkpo)
                    && OkpoRegex.IsMatch(creatorOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = creatorOkpo,
                Message = "Формат ввода данных не соответствует приказу. Укажите код ОКПО организации изготовителя."
            });
        }
        return result;
    }

    #endregion

    #region Check031

    //Если в коде ОКПО изготовителя указано примечание, то проверить его наличие (графа 8)
    private static List<CheckError> Check_031(List<Form12> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        string[] creatorOkpoValid = { "прим.", "прим", "примечание", "примечания" };
        if (!creatorOkpoValid.Contains(forms[line].CreatorOKPO_DB?.ToLower()) 
            && OKSM.All(oksmEntry => oksmEntry["shortname"] != forms[line].CreatorOKPO_DB)) return result;
        const byte graphNumber = 8;
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "CreatorOKPO_DB",
                Value = forms[line].CreatorOKPO_DB,
                Message = "Необходимо указать в примечании наименование и адрес организации-изготовителя ИОУ."
            });
        }
        return result;
    }

    #endregion

    #region Check032

    //Дата выпуска корректно заполнена (графа 9)
    private static List<CheckError> Check_032(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var creationDate = forms[line].CreationDate_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(creationDate) || creationDate is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "CreationDate_DB",
                Value = Convert.ToString(creationDate),
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        else if (!DateOnly.TryParse(creationDate, out _))
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "CreationDate_DB",
                Value = Convert.ToString(creationDate),
                Message = "Формат ввода данных не соответствует приказу. Некорректно заполнена дата выпуска."
            });
        }
        return result;
    }

    #endregion

    #region Check033

    //Дата выпуска менее даты операции
    private static List<CheckError> Check_033(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        if (!DateOnly.TryParse(forms[line].OperationDate_DB, out var opDate)
            || !DateOnly.TryParse(forms[line].CreationDate_DB, out var createDate))
        {
            return result;
        }
        var valid = createDate <= opDate;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "CreationDate_DB",
                Value = forms[line].CreationDate_DB,
                Message = "Дата выпуска не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check034

    //Проверить, истёк ли НСС
    private static List<CheckError> Check_034(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var signedServicePeriod = forms[line].SignedServicePeriod_DB ?? string.Empty;
        var creationDate = forms[line].CreationDate_DB ?? string.Empty;
        var operationDate = forms[line].OperationDate_DB ?? string.Empty;
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        if (opCode is "41"
            || !DateOnly.TryParse(creationDate, out var creationDateReal)
            || !DateOnly.TryParse(operationDate, out var operationDateReal)
            || !double.TryParse(signedServicePeriod.Replace('.', ','), out var signedServicePeriodReal))
        {
            return result;
        }
        creationDateReal = creationDateReal.AddMonths((int)signedServicePeriodReal);
        creationDateReal = creationDateReal.AddDays((int)Math.Round(30 * (signedServicePeriodReal % 1.0)));
        var valid = creationDateReal >= operationDateReal;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "SignedServicePeriod_DB",
                Value = signedServicePeriod,
                Message = "Для ИОУ истек НСС, следует продлить НСС либо снять с учета " +
                          "с одновременной постановкой на учет как РАО."
            });
        }
        return result;
    }

    #endregion

    #region Check035

    //Код формы собственности от 1 до 6, 9 (графа 11)
    private static List<CheckError> Check_035(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        if (propertyCode is not (>= 1 and <= 6 or 9))
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = propertyCode.ToString(),
                Message = "Формат ввода данных не соответствует приказу. " +
                          "Выберите идентификатор, соответствующий форме собственности ИОУ."
            });
        }
        return result;
    }

    #endregion

    #region Check036

    //Проверка наличия примечания
    private static List<CheckError> Check_036(List<Form12> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        const byte graphNumber = 11;
        byte?[] propertyCodeValid = { 2 };
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        var valid = !string.IsNullOrEmpty(owner)
                    && CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = propertyCode?.ToString(),
                Message = "Необходимо указать в примечании наименование субъекта Российской Федерации, " +
                          "в собственности которого находится объект учета."
            });
        }
        return result;
    }

    #endregion

    #region Check037

    //Проверка наличия примечания
    private static List<CheckError> Check_037(List<Form12> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        const byte graphNumber = 11;
        byte?[] propertyCodeValid = { 3 };
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        var valid = !string.IsNullOrEmpty(owner)
                    && CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = propertyCode?.ToString(),
                Message = "Необходимо указать в примечании наименование муниципального образования, " +
                          "в собственности которого находится объект учета."
            });
        }
        return result;
    }

    #endregion

    #region Check038

    //Проверка наличия примечания
    private static List<CheckError> Check_038(List<Form12> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        const byte graphNumber = 11;
        byte?[] propertyCodeValid = { 9 };
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        var valid = !string.IsNullOrEmpty(owner)
                    && CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PropertyCode_DB",
                Value = propertyCode?.ToString(),
                Message = "Необходимо указать в примечании наименование и адрес правообладателя " +
                          "(собственника или обладателя иного вещного права) на ИОУ."
            });
        }
        return result;
    }

    #endregion

    #region Check039

    //Код ОКПО правообладателя 8 или 14 цифр (графа 12)
    private static List<CheckError> Check_039(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (propertyCode is not (>= 1 and <= 4)) return result;
        var valid = !string.IsNullOrEmpty(owner)
                    && OkpoRegex.IsMatch(owner);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Формат ввода данных не соответствует приказу. " +
                          "В случае, если правообладатель - российское юридическое лицо, необходимо указать его код ОКПО."
            });
        }
        return result;
    }

    #endregion

    #region Check040

    //Код ОКПО правообладателя из ОКСМ (графа 12)
    private static List<CheckError> Check_040(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        byte?[] propertyCodeValid = { 5 };
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == owner)
                    && !owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Формат ввода данных не соответствует приказу. В случае, если правообладатель - иностранное государство, " +
                          "необходимо указать его краткое наименование в соответствии с ОКСМ."
            });
        }
        return result;
    }

    #endregion

    #region Check041

    //Код ОКПО правообладателя из ОКСМ, обязательно примечание (графа 12)
    private static List<CheckError> Check_041(List<Form12> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        byte?[] propertyCodeValid = { 6 };
        const byte graphNumber = 12;
        var propertyCode = forms[line].PropertyCode_DB;
        var owner = forms[line].Owner_DB;
        if (!propertyCodeValid.Contains(propertyCode)) return result;
        if (OKSM.All(oksmEntry => oksmEntry["shortname"] != owner)
            || owner.Equals("россия", StringComparison.CurrentCultureIgnoreCase))
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = owner,
                Message = "Необходимо указать в примечании наименование и адрес правообладателя " +
                          "(собственника или обладателя иного вещного права) на ИОУ."
            });
            return result;
        }
        var valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "Owner_DB",
                Value = forms[line].Owner_DB,
                Message = "Необходимо указать в примечании наименование и адрес правообладателя " +
                          "(собственника или обладателя иного вещного права) на ИОУ."
            });
        }
        return result;
    }

    #endregion

    #region Check042

    //Вид документа равен 1 
    private static List<CheckError> Check_042(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var documentVid = forms[line].DocumentVid_DB ?? 0;
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        if (opCode is "10" && documentVid is not 1)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = "Для операции инвентаризации сопровождающий документ - акт инвентаризации (код 1)."
            });
        }
        return result;
    }

    #endregion

    #region Check043

    //Вид документа равен 13
    private static List<CheckError> Check_043(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var documentVid = forms[line].DocumentVid_DB ?? 0;
        var opCode = forms[line].OperationCode_DB ?? "";
        if (opCode is "66" && documentVid is not 13)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = "Для операции продления НСС сопровождающий документ -  решение о продлении НСС (код 13)."
            });
        }
        return result;
    }

    #endregion

    #region Check044

    //Вид документа равен 1-15,19
    private static List<CheckError> Check_044(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var documentVid = forms[line].DocumentVid_DB ?? 0;
        var valid = documentVid is >= 1 and <= 15 or 19;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentVid_DB",
                Value = Convert.ToString(documentVid),
                Message = "Формат ввода данных не соответствует приказу. Необходимо указать вид документа в соответствии " +
                          "с таблицей 3 приложения №2 к приказу Госкорпорации \"Росатом\" от 07.12.2020 №1/13-НПА."
            });
        }
        return result;
    }

    #endregion

    #region Check045

    //Номер документа заполнен
    private static List<CheckError> Check_045(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var documentNumber = forms[line].DocumentNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(documentNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentNumber_DB",
                Value = documentNumber,
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check046

    //Дата документа корректно заполнена (графа 15)
    private static List<CheckError> Check_046(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var docDate = forms[line].DocumentDate_DB ?? string.Empty;
        if (string.IsNullOrWhiteSpace(docDate) || docDate is "-")
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
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
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(docDate),
                Message = "Формат ввода данных не соответствует приказу. Некорректно заполнена дата документа."
            });
        }
        return result;
    }

    #endregion

    #region Check047

    //Дата документа <= даты операции
    private static List<CheckError> Check_047(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] excludedOperationCodes = { "10", "41" };
        var operationCode = forms[line].OperationCode_DB;
        var operationDate = forms[line].OperationDate_DB;
        var documentDate = forms[line].DocumentDate_DB;
        if (excludedOperationCodes.Contains(operationCode)
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
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата документа не может быть позже даты операции."
            });
        }
        return result;
    }

    #endregion

    #region Check048

    //Дата операции равна дате документа
    private static List<CheckError> Check_048(List<Form12> forms, int line)
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
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = documentDate,
                Message = "Дата документа должна соответствовать дате операции."
            });
        }
        return result;
    }

    #endregion

    #region Check049

    //При коде операции 10, дата окончания ОП не позднее даты документа + 10 дней
    private static List<CheckError> Check_049(List<Form12> forms, Report rep, int line)
    {
        List<CheckError> result = new();
        var opCode = forms[line].OperationCode_DB ?? string.Empty;
        if (opCode is not "10" 
            || !DateOnly.TryParse(rep.EndPeriod_DB, out var pEnd)
            || !DateOnly.TryParse(forms[line].DocumentDate_DB, out var docDate))
        {
            return result;
        }
        var valid = WorkdaysBetweenDates(docDate, pEnd) <= 10;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "DocumentDate_DB",
                Value = Convert.ToString(forms[line].DocumentDate_DB),
                Message = "Дата документа не входит в отчетный период. Для операции инвентаризации, " +
                          "срок предоставления отчета исчисляется с даты утверждения акта инвентаризации."
            });
        }
        return result;
    }

    #endregion

    #region Check050

    //ОКПО поставщика/получателя равен коду ОКПО отчитывающейся организации (графа 16)
    private static List<CheckError> Check_050(List<Form12> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = 
        {
            "10", "11", "12", "15", "17", "18", "41", "42", "46", "58", "61", 
            "62", "67", "68", "71", "72", "73", "75", "97", "98", "99"
        };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var repOkpo = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;

        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo)
                     && providerOrRecieverOkpo == repOkpo;
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "Для выбранного кода операции указывается код ОКПО отчитывающейся организации."
            });
        }
        return result;
    }

    #endregion

    #region Check051

    //ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации (графа 16)
    private static List<CheckError> Check_051(List<Form12> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "25", "27", "28", "29", "35", "37", "38", "39" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var okpoRep = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        if (okpoRep == providerOrRecieverOkpo)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "Для выбранного кода операции указывается код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check052

    //ОКПО поставщика/получателя 8 или 14 цифр (графа 16)
    private static List<CheckError> Check_052(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21", "25", "27", "28", "29", "31", "35", "37", "38", "39", "54", "63", "64", "66"
        };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "Формат ввода данных не соответствует приказу. Укажите код ОКПО контрагента."
            });
        }
        return result;
    }

    #endregion

    #region Check053

    //ОКПО поставщика/получателя не равен коду ОКПО отчитывающейся организации (графа 16)
    private static List<CheckError> Check_053(List<Form12> forms, List<Form10> forms10, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "66" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        var okpoRep = !string.IsNullOrWhiteSpace(forms10[1].Okpo_DB)
            ? forms10[1].Okpo_DB
            : forms10[0].Okpo_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        if (okpoRep == providerOrRecieverOkpo)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "Для выбранного кода операции указывается код ОКПО организации, осуществившей продление НСС."
            });
        }
        return result;
    }

    #endregion

    #region Check054

    //ОКПО поставщика/получателя 8 или 14 цифр, или "Минобороны" (графа 16)
    private static List<CheckError> Check_054(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes = { "22", "32" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(providerOrRecieverOkpo)
                    || providerOrRecieverOkpo.Equals("минобороны", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "Формат ввода данных не соответствует приказу. " +
                          "Следует указать код ОКПО контрагента, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check055

    //ОКПО поставщика/получателя из ОКСМ (графа 16)
    private static List<CheckError> Check_055(List<Form12> forms, List<Note> notes, int line)
    {
        List<CheckError> result = new();
        const byte graphNumber = 16;
        string[] applicableOperationCodes = { "81", "82", "83", "84", "85", "86", "87", "88" };
        var operationCode = forms[line].OperationCode_DB;
        var providerOrRecieverOkpo = forms[line].ProviderOrRecieverOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OKSM.Any(oksmEntry => oksmEntry["shortname"] == providerOrRecieverOkpo)
                    && !providerOrRecieverOkpo.Equals("россия", StringComparison.CurrentCultureIgnoreCase);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "Формат ввода данных не соответствует приказу."
            });
        }
        valid = CheckNotePresence(notes, line, graphNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "ProviderOrRecieverOKPO_DB",
                Value = providerOrRecieverOkpo,
                Message = "Необходимо добавить примечание."
            });
        }
        return result;
    }

    #endregion

    #region Check056

    //Код ОКПО перевозчика равен "-" (графа 17)
    private static List<CheckError> Check_056(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "10", "11", "12", "15", "17", "18", "41", "42", "43", "46", "53", "54", 
            "58", "66", "67", "68", "71", "72", "73", "74", "75", "97", "98"
        };
        if (!applicableOperationCodes.Contains(forms[line].OperationCode_DB)) return result;
        var transporterOkpo = forms[line].TransporterOKPO_DB;
        var valid = transporterOkpo is "-";
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = transporterOkpo,
                Message = "При выбранном коде операции транспортирование не производится."
            });
        }
        return result;
    }

    #endregion

    #region Check057

    //Код ОКПО перевозчика равен 8/14 цифр (графа 17)
    private static List<CheckError> Check_057(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        string[] applicableOperationCodes =
        {
            "21", "25", "27", "28", "29", "31", "35", "36", "37", "38", "39", 
            "61", "62", "81", "82", "83", "84", "85", "86", "87", "88"
        };
        var operationCode = forms[line].OperationCode_DB;
        var transporterOkpo = forms[line].TransporterOKPO_DB;
        if (!applicableOperationCodes.Contains(operationCode)) return result;
        var valid = OkpoRegex.IsMatch(transporterOkpo);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = transporterOkpo,
                Message = "Необходимо указать код ОКПО организации перевозчика."
            });
        }
        return result;
    }

    #endregion

    #region Check058

    //Код ОКПО перевозчика состоит из 8/14 чисел или "минобороны" (графа 17)
    private static List<CheckError> Check_058(List<Form12> forms, int line)
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
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "TransporterOKPO_DB",
                Value = transporterOkpo,
                Message = "Необходимо указать код ОКПО организации перевозчика, либо \"Минобороны\" без кавычек."
            });
        }
        return result;
    }

    #endregion

    #region Check059

    //Наименование прибора заполнено (графа 21)
    private static List<CheckError> Check_059(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var packName = forms[line].PackName_DB;
        var valid = !string.IsNullOrWhiteSpace(packName);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PackName_DB",
                Value = packName,
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check060

    //Тип УКТ заполнен (графа 22)
    private static List<CheckError> Check_060(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var packType = forms[line].PackType_DB;
        var valid = !string.IsNullOrWhiteSpace(packType);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PackType_DB",
                Value = packType,
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #region Check061

    //Номер УКТ заполнен (графа 23)
    private static List<CheckError> Check_061(List<Form12> forms, int line)
    {
        List<CheckError> result = new();
        var packNumber = forms[line].PackNumber_DB;
        var valid = !string.IsNullOrWhiteSpace(packNumber);
        if (!valid)
        {
            result.Add(new CheckError
            {
                FormNum = "form_12",
                Row = (line + 1).ToString(),
                Column = "PackNumber_DB",
                Value = packNumber,
                Message = "Формат ввода данных не соответствует приказу. Графа не может быть пустой."
            });
        }
        return result;
    }

    #endregion

    #endregion

    #region GraphsList

    private static Dictionary<string, string> GraphsList { get; } = new()
    {
        { "NumberInOrder_DB", "01 - № п/п" },
        { "OperationCode_DB", "02 - Код операции" },
        { "OperationDate_DB", "03 - Дата операции" },
        { "PassportNumber_DB", "04 - Номер паспорта (сертификата)" },
        { "NameIOU_DB", "05 - Наименование ИОУ" },
        { "FactoryNumber_DB", "06 - Номер" },
        { "Mass_DB", "07 - Масса ОУ, кг." },
        { "CreatorOKPO_DB", "08 - Код ОКПО изготовителя" },
        { "CreationDate_DB", "09 - Дата выпуска" },
        { "SignedServicePeriod_DB", "10 - НСС, месяцев" },
        { "PropertyCode_DB", "11 - Код формы собственности" },
        { "Owner_DB", "12 - Код ОКПО правообладателя" },
        { "DocumentVid_DB", "13 - Документ - вид" },
        { "DocumentNumber_DB", "14 - Документ - номер" },
        { "DocumentDate_DB", "15 - Документ - дата" },
        { "ProviderOrRecieverOKPO_DB", "16 - Код ОКПО поставщика или получателя" },
        { "TransporterOKPO_DB", "17 - Код ОКПО перевозчика" },
        { "PackName_DB", "18 - Наименование прибора (упаковки), УКТ или иной упаковки" },
        { "PackType_DB", "19 - Тип прибора (упаковки), УКТ или иной упаковки" },
        { "PackNumber_DB", "20 - Номер прибора (упаковки), УКТ или иной упаковки" }
    };

    #endregion
}
