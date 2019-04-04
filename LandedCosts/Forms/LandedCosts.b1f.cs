using System;
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
            EditText0 = ((EditText)(GetItem("Item_0").Specific));
            EditText0.ChooseFromListAfter += new _IEditTextEvents_ChooseFromListAfterEventHandler(EditText0_ChooseFromListAfter);
            StaticText0 = ((StaticText)(GetItem("Item_1").Specific));
            StaticText1 = ((StaticText)(GetItem("Item_2").Specific));
            EditText1 = ((EditText)(GetItem("Item_3").Specific));
            EditText1.ChooseFromListAfter += new _IEditTextEvents_ChooseFromListAfterEventHandler(EditText1_ChooseFromListAfter);
            EditText2 = ((EditText)(GetItem("Item_4").Specific));
            StaticText2 = ((StaticText)(GetItem("Item_5").Specific));
            EditText3 = ((EditText)(GetItem("Item_6").Specific));
            StaticText3 = ((StaticText)(GetItem("Item_7").Specific));
            Grid0 = ((Grid)(GetItem("Item_8").Specific));
            Grid0.DoubleClickAfter += new _IGridEvents_DoubleClickAfterEventHandler(Grid0_DoubleClickAfter);
            Grid0.ClickAfter += new _IGridEvents_ClickAfterEventHandler(Grid0_ClickAfter);
            Button0 = ((Button)(GetItem("Item_9").Specific));
            Button0.PressedAfter += new _IButtonEvents_PressedAfterEventHandler(Button0_PressedAfter);
            OnCustomInitialize();

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
                _cardCode = ((EditText)bpMatrix.Columns.Item(1).Cells.Item(selectedRow).Specific).Value;
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
                    _invoiceDocEntry = ((EditText)bpMatrix.Columns.Item(1).Cells.Item(selectedRow).Specific).Value;
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
                CreateChoseFromList(Application.SBO_Application.Forms.ActiveForm, "2", "CFL11", "Item_0", "CardCode", "DT_1","CardType", BoConditionOperation.co_EQUAL, "S",  "CFL2");

   
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
        private static void CreateChoseFromList(Form form, string objectType, string cflId, string editTextItem, string cflAlias,  string dataSourceId, string conditionField = "", BoConditionOperation equality = BoConditionOperation.co_EQUAL, string conditionValue = "", string cflParamsId = "")
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
                    InvoiceDocEntry = int.Parse(EditText1.Value),
                    PostingDate = DateTime.ParseExact(EditText2.Value, "yyyyMMdd", CultureInfo.InvariantCulture),
                    Number = EditText3.Value,
                    VendorCode = EditText0.Value
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

                Application.SBO_Application.StatusBar.SetSystemMessage("წარმატება", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
                var res = invoice.Update();

                landedCost.Series = 23;


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
                    LandedCost_ItemLine oLandedCostItemLine = landedCost.LandedCost_ItemLines.Add();
                    oLandedCostItemLine.BaseDocumentType = LandedCostBaseDocumentTypeEnum.asPurchaseInvoice;
                    oLandedCostItemLine.BaseEntry = model.InvoiceDocEntry;
                    oLandedCostItemLine.BaseLine = i;
                }



                LandedCost_CostLine costLine = landedCost.LandedCost_CostLines.Add();
                costLine.LandedCostCode = model.Rows[0].LandedCostCode;
                if (landedCost.DocumentCurrency != "GEL")
                {
                    costLine.AmountFC = model.Rows[0].Price;
                }
                else
                {
                    costLine.amount = model.Rows[0].Price;
                }

                landedCost.Broker = model.Rows[0].Broker;

                try
                {

                    landedCost.PostingDate = model.PostingDate;
                    LandedCostParams landedCostParams = svrLandedCost.AddLandedCost(landedCost);
                    landedCostParams.LandedCostNumber = landedCostParams.LandedCostNumber;
                    var addedLandedCost = svrLandedCost.GetLandedCost(landedCostParams);
                    //addedLandedCost.JournalRemarks = "gocha";
                    //svrLandedCost.UpdateLandedCost(addedLandedCost);
                    bool added = CreateInvoiceFromLandedCost(addedLandedCost, model);
                    //if (!added)
                    //{
                    //    if (DiManager.Company.InTransaction)
                    //    {
                    //        DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                    //    }
                    //    return false;
                    //}

                }
                catch (Exception e)
                {
                    if (DiManager.Company.InTransaction)
                    {
                        DiManager.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                    }
                    Application.SBO_Application.StatusBar.SetSystemMessage(e.Message, BoMessageTime.bmt_Short);
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
            invoice.Comments = addedLandedCost.DocEntry.ToString();

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
            invoice.Lines.LineTotal = model.Rows[0].Currency == "GEL" ? costLine.amount : costLine.AmountFC;

            int res = invoice.Add();
            Application.SBO_Application.SetStatusBarMessage(DiManager.Company.GetLastErrorDescription(), BoMessageTime.bmt_Short,
                true);
            return res == 0;
        }
    }
}