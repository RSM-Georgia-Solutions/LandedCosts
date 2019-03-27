using System;
using System.Collections.Generic;
using System.Xml;
using LandedCosts.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using ChooseFromList = SAPbouiCOM.ChooseFromList;
using Currencies = LandedCosts.Forms.Currencies;

namespace LandedCosts
{
    [FormAttribute("LandedCosts.LandedCosts", "Forms/LandedCosts.b1f")]
    class LandedCosts : UserFormBase
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
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ActivateAfter += new SAPbouiCOM.Framework.FormBase.ActivateAfterHandler(this.Form_ActivateAfter);
            this.VisibleAfter += new VisibleAfterHandler(this.Form_VisibleAfter);

        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {
            LandedCostsSetup.LandedCostsCode += FillLandedCostCode;
            Currencies.CurrencyCode += CurrencyCode;
            TaxGroups._taxCode += TaxCode;

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

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.Grid Grid0;


        private string _cardCode = string.Empty;
        private string _invoiceDocEntry = string.Empty;
        private void EditText0_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            Form listOfBps = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
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

        }

        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {
            EditText0.Value = _cardCode;
            EditText1.Value = _invoiceDocEntry;
        }

        private void EditText1_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            Form listOfBps = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (listOfBps.Title == "Landed Costs")
            {
                EditText1.Value = _invoiceDocEntry;
            }
            if (listOfBps.TypeEx == "10017")
            {
                Matrix bpMatrix = (Matrix)listOfBps.Items.Item("7").Specific;
                int selectedRow = bpMatrix.GetNextSelectedRow();
                _invoiceDocEntry = ((EditText)bpMatrix.Columns.Item(1).Cells.Item(selectedRow).Specific).Value;
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

        private void Grid0_DoubleClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ColUID == "დამატებითი ხარჯი" || pVal.ColUID == "დამატებითი ხარჯი სახელი")
            {
                _pvalRowDoubleCost = pVal.Row;
                SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("8456");
            }
            if (pVal.ColUID == "ვალუტა")
            {
                _pvalRowDoubleCur = pVal.Row;
                SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("8450");
            }
            if (pVal.ColUID == "დღგ-ს ჯგუფი")
            {
                _pvalRowDoubleTax = pVal.Row;
                SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("8458");
            }
            if (pVal.ColUID == "ბროკერი")
            {
                _pvalRowDoubleTax = pVal.Row;
                SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("8458");
            }
        }

        private void Form_VisibleAfter(SBOItemEventArg pVal)
        {
            try
            {
                var activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                ChooseFromListCollection oCfLs = activeForm.ChooseFromLists;
                ChooseFromListCreationParams oCflCreationParams = ((ChooseFromListCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCflCreationParams.MultiSelection = false;
                oCflCreationParams.ObjectType = "2";
                oCflCreationParams.UniqueID = "CFL1";
                ChooseFromList oCfl = oCfLs.Add(oCflCreationParams);
                EditText oEdit = (EditText)activeForm.Items.Item("Item_4").Specific;
                oEdit.DataBind.SetBound(true, "", "CardCode");
                oEdit.ChooseFromListAlias = "CardCode";
                Conditions oConditions = oCfl.GetConditions();
                Condition oCondition = oConditions.Add();
                oCondition.Alias = "CardType";
                oCondition.Operation = BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "S";
                oCfl.SetConditions(oConditions);
            }
            catch (Exception e)
            {
            }
        }

        private Button Button0;

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

            LandedCostsModel model = new LandedCostsModel();
            model.InvoiceDocEntry = int.Parse(EditText1.Value);
            model.PostingDate = DateTime.Parse(EditText2.Value);
            model.Number = EditText3.Value;

            for (int i = 0; i < Grid0.DataTable.Rows.Count; i++)
            {
                LandedCostsRowModel row = new LandedCostsRowModel();
            }



            int landedCostDocEntry;
            LandedCostsService svrLandedCost = (LandedCostsService)DiManager.Company.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService);
            SAPbobsCOM.LandedCost landedCost = (LandedCost)svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCost);


            Documents invoice = (SAPbobsCOM.Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
            invoice.GetByKey(int.Parse(EditText1.Value));
            invoice.OpenForLandedCosts = BoYesNoEnum.tYES;
            var res = invoice.Update();

            landedCost.Series = 23;

            for (int i = 0; i < invoice.Lines.Count; i++)
            {
                LandedCost_ItemLine oLandedCostItemLine = landedCost.LandedCost_ItemLines.Add();
                oLandedCostItemLine.BaseDocumentType = LandedCostBaseDocumentTypeEnum.asPurchaseInvoice;
                oLandedCostItemLine.BaseEntry = int.Parse(EditText1.Value);
                oLandedCostItemLine.BaseLine = i;
                SAPbobsCOM.LandedCost_CostLine costLine = landedCost.LandedCost_CostLines.Add();
                costLine.LandedCostCode = "01";
                costLine.amount = 100;
            }

            SAPbobsCOM.LandedCostParams landedCostParams = (LandedCostParams)svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams);

            try
            {
                landedCostParams = svrLandedCost.AddLandedCost(landedCost);
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetSystemMessage(e.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            landedCostDocEntry = landedCostParams.LandedCostNumber;
        }
    }
}