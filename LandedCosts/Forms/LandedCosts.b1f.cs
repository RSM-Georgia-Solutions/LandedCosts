﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Xml;
using LandedCosts.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;
using ChooseFromList = SAPbouiCOM.ChooseFromList;
using Currencies = LandedCosts.Forms.Currencies;

namespace LandedCosts
{
    [FormAttribute("LandedCosts.LandedCosts", "Forms/LandedCosts.b1f")]
    class LandedCosts : UserFormBase, IChoseFromList
    {
        public LandedCosts()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText1_ChooseFromListAfter);
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_8").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_10").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_11").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            ActivateAfter += new ActivateAfterHandler(Form_ActivateAfter);
            VisibleAfter += new VisibleAfterHandler(Form_VisibleAfter);

        }

        private EditText EditText0;

        private void OnCustomInitialize()
        {
            LandedCostsSetup.LandedCostsCode = FillLandedCostCode;
            Currencies.CurrencyCode = CurrencyCode;
            TaxGroups._taxCode = TaxCode;

            StaticText0.Item.FontSize = 15;
            StaticText0.Item.Height = 20;
            StaticText1.Item.FontSize = 15;
            StaticText1.Item.Height = 20;
            StaticText2.Item.FontSize = 15;
            StaticText2.Item.Height = 20;
            StaticText3.Item.FontSize = 15;
            StaticText3.Item.Height = 20;
            Button0.Item.FontSize = 15;
            Button0.Item.Height = 20;
            Button1.Item.FontSize = 15;
            Button1.Item.Height = 20;

            Grid0.DataTable.Columns.Add("დამატებითი ხარჯი", BoFieldsType.ft_AlphaNumeric);
            Grid0.DataTable.Columns.Add("დამატებითი ხარჯი სახელი", BoFieldsType.ft_AlphaNumeric);
            Grid0.DataTable.Columns.Add("თანხა", BoFieldsType.ft_Price);
            Grid0.DataTable.Columns.Add("ვალუტა", BoFieldsType.ft_AlphaNumeric);
            Grid0.DataTable.Columns.Add("დღგ-ს ჯგუფი", BoFieldsType.ft_AlphaNumeric);
            Grid0.DataTable.Columns.Add("ბროკერი", BoFieldsType.ft_AlphaNumeric);
            Grid0.DataTable.Columns.Add("კურსი", BoFieldsType.ft_Rate);
            Grid0.DataTable.Rows.Add();
            Grid0.AutoResizeColumns();
        }

        private void TaxCode(string taxCode)
        {
            if (_pvalRowDoubleTax == -1)
            {
                return;
            }
            Grid0.DataTable.SetValue("დღგ-ს ჯგუფი", _pvalRowDoubleTax, taxCode);
        }

        private void CurrencyCode(string curCode)
        {
            if (_pvalRowDoubleCur == -1)
            {
                return;
            }
            Grid0.DataTable.SetValue(3, _pvalRowDoubleCur, curCode);
        }

        private void FillLandedCostCode(string code, string name)
        {
            if (_pvalRowDoubleCost == -1)
            {
                return;
            }
            Grid0.DataTable.SetValue(0, _pvalRowDoubleCost, code);
            Grid0.DataTable.SetValue(1, _pvalRowDoubleCost, name);
        }

        private StaticText StaticText0;
        private StaticText StaticText1;
        private EditText EditText1;
        private EditText EditText2;
        private StaticText StaticText2;
        private EditText EditText3;
        private StaticText StaticText3;
        private Grid Grid0;


        private string _cardCode = string.Empty;
        private string _invoiceDocEntry = string.Empty;
        private string _invoiceDocNum = string.Empty;
        private void EditText0_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            Form listOfBps = Application.SBO_Application.Forms.ActiveForm;
            if (listOfBps.Title == "Landed Costs")
            {
                EditText0.Value = _cardCode;
            }
            if (listOfBps.TypeEx == "10001")
            {
                Matrix bpMatrix = (Matrix)listOfBps.Items.Item("7").Specific;
                int selectedRow = bpMatrix.GetNextSelectedRow();
                _cardCode = ((EditText)bpMatrix.Columns.Item("CardCode").Cells.Item(selectedRow).Specific).Value;
            }

            if (listOfBps.Title == "Landed Costs")
            {

                try
                {
                    CreateChoseFromList(Application.SBO_Application.Forms.ActiveForm, "18", "CFL13", "Item_3", "DocEntry", "DT_2", "CardCode", BoConditionOperation.co_EQUAL, _cardCode, "CFL3");
                }
                catch (Exception)
                {
                    Conditions oCons = Application.SBO_Application.Forms.ActiveForm.ChooseFromLists.Item("CFL13").GetConditions();
                    Condition oCon = oCons.Item(0);
                    oCon.CondVal = _cardCode;

                    Application.SBO_Application.Forms.ActiveForm.ChooseFromLists.Item("CFL13").SetConditions(oCons);
                    Application.SBO_Application.Forms.ActiveForm.ChooseFromLists.Item("CFL13").GetConditions().Item(0).CondVal = _cardCode;
                }

            }


        }

        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {
            EditText0.Value = _cardCode;
            EditText1.Value = _invoiceDocEntry;         
        }

        private void EditText1_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            Form listOfBps = Application.SBO_Application.Forms.ActiveForm;
            if (listOfBps.Title == "Landed Costs")
            {
                EditText1.Value = _invoiceDocEntry;
            }
            if (listOfBps.TypeEx == "10017")
            {
                Matrix bpMatrix = (Matrix)listOfBps.Items.Item("7").Specific;
                int selectedRow = bpMatrix.GetNextSelectedRow();
                try
                {
                    _invoiceDocNum = ((EditText)bpMatrix.Columns.Item("DocNum").Cells.Item(selectedRow).Specific).Value;
                    _invoiceDocEntry = ((EditText)bpMatrix.Columns.Item("V_0").Cells.Item(selectedRow).Specific).Value;//DocEntry
                }
                catch (Exception e)
                {

                }
            }
        }

        private void Grid0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.Row + 1 == Grid0.DataTable.Rows.Count)
            {
                Grid0.DataTable.Rows.Add();
            }
        }


        private int _pvalRowDoubleCost = 0;
        private int _pvalRowDoubleCur = 0;
        private int _pvalRowDoubleVat = 0;
        private int _pvalRowDoubleTax = 0;
        private int _pvalRowDoubleBroker = 0;

        public string _bpCardCode = string.Empty;
        public string _bpName = string.Empty;

        public void AfterChooseFromList(string bpCode, string bpName)
        {
            Grid0.DataTable.SetValue(5, _pvalRowDoubleBroker, bpCode);
        }

        private void Grid0_DoubleClickAfter(object sboObject, SBOItemEventArg pVal)
        {

            if (pVal.ColUID == "დამატებითი ხარჯი" || pVal.ColUID == "დამატებითი ხარჯი სახელი")
            {
                _pvalRowDoubleCost = pVal.Row;
                Application.SBO_Application.ActivateMenuItem("8456");
            }
            if (pVal.ColUID == "ვალუტა")
            {
                _pvalRowDoubleCur = pVal.Row;
                Application.SBO_Application.ActivateMenuItem("8450");
            }
            if (pVal.ColUID == "დღგ-ს ჯგუფი")
            {
                _pvalRowDoubleTax = pVal.Row;
                Application.SBO_Application.ActivateMenuItem("8458");
            }
            if (pVal.ColUID == "ბროკერი")
            {
                _pvalRowDoubleBroker = pVal.Row;
                ListOfBusinessPartners businessPartners = new ListOfBusinessPartners(this);
                businessPartners.Show();
            }
        }

        private void Form_VisibleAfter(SBOItemEventArg pVal)
        {
            try
            {
                CreateChoseFromList(Application.SBO_Application.Forms.ActiveForm, "2", "CFL11", "Item_0", "CardCode", "DT_1", "CardType", BoConditionOperation.co_EQUAL, "S", "CFL2");


            }
            catch (Exception e)
            {
            }
        }
        /// <summary>
        /// objectType - choseFromList Object
        /// cflId - chooseFromList Uniq Id
        /// editTextItem - editText Item for CFL
        /// conditionField - Filter for CFL Field
        /// cflAlias - CFL Field
        /// </summary>
        /// <param name="form"></param>
        /// <param name="objectType"></param>
        /// <param name="cflId"></param>
        /// <param name="editTextItem"></param>
        /// <param name="conditionField"></param>
        /// <param name="equality"></param>
        /// <param name="conditionValue"></param>
        /// <param name="cflAlias"></param>
        /// <param name="cflParamsId"></param>
        /// <param name="dataSourceId"></param>
        private static void CreateChoseFromList(Form form, string objectType, string cflId, string editTextItem, string cflAlias, string dataSourceId, string conditionField = "", BoConditionOperation equality = BoConditionOperation.co_EQUAL, string conditionValue = "", string cflParamsId = "")
        {
            Form activeForm = form;
            ChooseFromListCollection oCfLs = activeForm.ChooseFromLists;
            ChooseFromListCreationParams oCflCreationParams = (ChooseFromListCreationParams)Application.SBO_Application.CreateObject(BoCreatableObjectType
                .cot_ChooseFromListCreationParams);
            oCflCreationParams.ObjectType = objectType;
            oCflCreationParams.UniqueID = cflId;
            ChooseFromList oCfl = oCfLs.Add(oCflCreationParams);
            Conditions oCons = oCfl.GetConditions();
            Condition oCon = oCons.Add();

            if (conditionField != "")
            {
                oCon.Alias = conditionField;
                oCon.Operation = equality;
                oCon.CondVal = conditionValue;
                oCfl.SetConditions(oCons);
            }
            oCflCreationParams.UniqueID = cflParamsId;
            oCfl = oCfLs.Add(oCflCreationParams);
            activeForm.DataSources.UserDataSources.Add(dataSourceId, BoDataType.dt_SHORT_TEXT, 254);
            EditText oEdit = ((EditText)(activeForm.Items.Item(editTextItem).Specific));
            oEdit.DataBind.SetBound(true, "", dataSourceId);
            oEdit.ChooseFromListUID = cflId;
            oEdit.ChooseFromListAlias = cflAlias;
        }

        private Button Button0;

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            List<LandedCostsModel> modelsList = new List<LandedCostsModel>();
            if (string.IsNullOrWhiteSpace(EditText2.Value) || string.IsNullOrWhiteSpace(EditText1.Value) || string.IsNullOrWhiteSpace(EditText3.Value) || string.IsNullOrWhiteSpace(EditText0.Value))
            {
                Application.SBO_Application.SetStatusBarMessage("შეავსეთ ველები",
                    BoMessageTime.bmt_Short, true);
                return;
            }


            for (int i = 0; i < Grid0.DataTable.Rows.Count - 1; i++)
            {
                LandedCostsModel model = new LandedCostsModel
                {
                    InvoiceDocNum = int.Parse(EditText1.Value),
                    InvoiceDocEntry = int.Parse(_invoiceDocEntry),
                    PostingDate = DateTime.ParseExact(EditText2.Value, "yyyyMMdd", CultureInfo.InvariantCulture),
                    Number = EditText3.Value,
                    VendorCode = EditText0.Value,
                    Comment = EditText4.Value
                };

                LandedCostsRowModel row = new LandedCostsRowModel
                {
                    Broker = Grid0.DataTable.GetValue(5, i).ToString(),
                    Currency = Grid0.DataTable.GetValue(3, i).ToString(),
                    LandedCostCode = Grid0.DataTable.GetValue(0, i).ToString(),
                    LandedCostName = Grid0.DataTable.GetValue(1, i).ToString(),
                    Price = double.Parse(Grid0.DataTable.GetValue(2, i).ToString()),
                    Rate = double.Parse(Grid0.DataTable.GetValue(6, i).ToString()),
                    VatGroup = Grid0.DataTable.GetValue(4, i).ToString()
                };
                model.Rows.Add(row);
                modelsList.Add(model);
            }

            DiManager.Company.StartTransaction();

            if (modelsList.Count < 1)
            {
                DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                return;
            }

            bool success = PostLandedCosts(modelsList);

            if (success)
            {
                Application.SBO_Application.MessageBox("წარმატება");
                if (DiManager.Company.InTransaction)
                {
                    DiManager.Company.EndTransaction(BoWfTransOpt.wf_Commit);
                }
            }
        }

        private static bool PostLandedCosts(List<LandedCostsModel> modelsList)
        {
            foreach (var model in modelsList)
            {
                if (model.Rows[0].LandedCostCode == "")
                {
                    continue;
                }
                LandedCostsService svrLandedCost = (LandedCostsService)DiManager.Company.GetCompanyService().GetBusinessService(ServiceTypes.LandedCostsService);
                LandedCost landedCost = (LandedCost)svrLandedCost.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost);
                Documents invoice = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
                invoice.GetByKey(model.InvoiceDocEntry);
                invoice.OpenForLandedCosts = BoYesNoEnum.tYES;
                try
                {
                    invoice.JournalMemo = "";
                    invoice.JournalMemo = invoice.JournalMemo + $"D: { model.Number}";
                }
                catch (Exception e)
                {
                }

                var res = invoice.Update();

                landedCost.Series = 23;
                landedCost.Remarks = model.Comment;


                try
                {
                    landedCost.DocumentCurrency = model.Rows[0].Currency;
                    landedCost.DocumentRate = model.Rows[0].Rate;
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                    // cant change currency 
                }

                for (int i = 0; i < invoice.Lines.Count; i++)
                {

                    invoice.Lines.SetCurrentLine(i);
                    var lineNum = invoice.Lines.LineNum;

                    LandedCost_ItemLine oLandedCostItemLine = landedCost.LandedCost_ItemLines.Add();
                    oLandedCostItemLine.BaseDocumentType = LandedCostBaseDocumentTypeEnum.asPurchaseInvoice;
                    oLandedCostItemLine.BaseEntry = model.InvoiceDocEntry;
                    oLandedCostItemLine.BaseLine = lineNum;
                }


                Recordset recSet =
                    (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes
                        .BoRecordset);

                recSet.DoQuery(DiManager.QueryHanaTransalte($"SELECT Rate FROM OVTG WHERE Code = N'{model.Rows[0].VatGroup}'"));

                double vatPercent = -1;

                try
                {
                    vatPercent = double.Parse(recSet.Fields.Item("Rate").Value.ToString(), CultureInfo.InvariantCulture);
                }
                catch (Exception e)
                {
                    Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                }


                LandedCost_CostLine costLine = landedCost.LandedCost_CostLines.Add();
                costLine.LandedCostCode = model.Rows[0].LandedCostCode;

                if (landedCost.DocumentCurrency != "GEL")
                {
                    if (vatPercent == 18)
                    {
                        costLine.AmountFC = Math.Round(model.Rows[0].Price / 1.18, 2);
                    }
                    else
                    {
                        costLine.AmountFC = model.Rows[0].Price;
                    }
                }
                else
                {
                    if (vatPercent == 18)
                    {
                        costLine.amount = Math.Round(model.Rows[0].Price / 1.18, 2);
                    }
                    else
                    {
                        costLine.amount = Math.Round(model.Rows[0].Price, 2);
                    }
                }

                landedCost.Broker = model.Rows[0].Broker;

                try
                {

                    landedCost.PostingDate = model.PostingDate;
                    LandedCostParams landedCostParams = svrLandedCost.AddLandedCost(landedCost);
                    landedCostParams.LandedCostNumber = landedCostParams.LandedCostNumber;
                    var addedLandedCost = svrLandedCost.GetLandedCost(landedCostParams);
                    bool added = CreateInvoiceFromLandedCost(addedLandedCost, model);

                    if (!added)
                    {
                        if (DiManager.Company.InTransaction)
                        {
                            DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);

                        }
                        return false;
                    }

                }
                catch (Exception e)
                {
                    if (DiManager.Company.InTransaction)
                    {
                        DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                    }
                    Application.SBO_Application.StatusBar.SetSystemMessage("Landed Cost " + e.Message, BoMessageTime.bmt_Short);
                    return false;
                }
            }
            return true;
        }

        private static bool CreateInvoiceFromLandedCost(LandedCost addedLandedCost, LandedCostsModel model)
        {
            Documents invoice = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
            invoice.CardCode = addedLandedCost.Broker;
            invoice.DocDate = addedLandedCost.PostingDate;
            invoice.VatDate = addedLandedCost.PostingDate;
            invoice.DocType = BoDocumentTypes.dDocument_Service;

            invoice.DocCurrency = model.Rows[0].Currency;
            invoice.DocRate = model.Rows[0].Rate;
            invoice.Comments = $"Landed Cost : {addedLandedCost.DocEntry} Invoice(s) : {model.Comment} ";
            invoice.JournalMemo = $"Declaration : {model.Number}";

            var costLine = addedLandedCost.LandedCost_CostLines.Item(0);
            var landedCostCode = costLine.LandedCostCode;
            var jdtCode = addedLandedCost.TransactionNumber;

            Recordset recSet =
                (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes
                    .BoRecordset);
            recSet.DoQuery(DiManager.QueryHanaTransalte($"SELECT * FROM OALC where AlcCode = '{landedCostCode}'"));
            var landedCostName = recSet.Fields.Item("AlcName").Value.ToString();
            var landedCostAccount = recSet.Fields.Item("LaCAllcAcc").Value.ToString();

            recSet.DoQuery(DiManager.QueryHanaTransalte($"SELECT BplId FROM JDT1 WHERE transId =  '{jdtCode}'"));
            int branch = int.Parse(recSet.Fields.Item("BplId").Value.ToString());
            invoice.BPL_IDAssignedToInvoice = branch;

            //Invoice Lines
            invoice.Lines.SetCurrentLine(0);
            invoice.Lines.ItemDescription = landedCostName;
            invoice.Lines.VatGroup = model.Rows[0].VatGroup;
            invoice.Lines.AccountCode = landedCostAccount;
            invoice.Lines.PriceAfterVAT = model.Rows[0].Price;

            int res = invoice.Add();
            if (res == 0)
            {
                return true;
            }
            Application.SBO_Application.SetStatusBarMessage("Invoice " + DiManager.Company.GetLastErrorDescription(), BoMessageTime.bmt_Short,
                true);
            return false;
        }

        private EditText EditText4;
        private Button Button1;

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            Application.SBO_Application.ActivateMenuItem("1540");
            if (Application.SBO_Application.Forms.ActiveForm.Type != 392) return;
            Form journalEntry = Application.SBO_Application.Forms.ActiveForm;
            try
            {
                ((ComboBox)journalEntry.Items.Item("9").Specific).Select("01");
            }
            catch (Exception)
            {
                Application.SBO_Application.SetStatusBarMessage("ტრანზაქციის კოდი არ არსებობს",
                    BoMessageTime.bmt_Short, true);
            }
            ((EditText)journalEntry.Items.Item("6").Specific).Value = EditText2.Value;
            ((EditText)journalEntry.Items.Item("102").Specific).Value = EditText2.Value;
            ((EditText)journalEntry.Items.Item("97").Specific).Value = EditText2.Value;
            ((EditText)journalEntry.Items.Item("1000").Specific).Value = EditText2.Value;
            ((EditText)journalEntry.Items.Item("10").Specific).Value = EditText3.Value;
            ((EditText)journalEntry.Items.Item("7").Specific).Value = $"Declaration : {EditText3.Value}";
        }

    }
}