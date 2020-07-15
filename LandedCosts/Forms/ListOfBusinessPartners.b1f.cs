using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace LandedCosts.Forms
{
    [FormAttribute("LandedCosts.Forms.ListOfBusinessPartners", "Forms/ListOfBusinessPartners.b1f")]
    class ListOfBusinessPartners : UserFormBase
    {
        private IChoseFromList _landed;
        public ListOfBusinessPartners(IChoseFromList cost)
        {
            _landed = cost;
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_3").Specific));
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte("SELECT CardName as 'BP Name', CardCode as 'BP Code', Balance, VatIdUnCmp as 'Unified Federal Tax ID', GroupCode as 'Group Code', DebPayAcct as 'Account' FROM OCRD WHERE OCRD.CardType = 'S' "));
            EditTextColumn oEditCol = (EditTextColumn)Grid0.Columns.Item("BP Code");
            oEditCol.LinkedObjectType = "2";
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;




        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            var selectedRow = Grid0.Rows.SelectedRows;
            if (selectedRow.Count == 0)
            {
                return;
            }
            var z = selectedRow.Item(0, BoOrderType.ot_RowOrder);

            _landed.AfterChooseFromList(Grid0.DataTable.GetValue("BP Code", z).ToString(), Grid0.DataTable.GetValue("BP Name", z).ToString());
            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
        }

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte("SELECT CardName as 'BP Name', CardCode as 'BP Code', Balance, VatIdUnCmp as 'Unified Federal Tax ID', GroupCode as 'Group Code', DebPayAcct as 'Account' FROM OCRD WHERE (OCRD.CardType = 'S' and CardName = N'" + EditText0.Value + "') or  (OCRD.CardType = 'S' and CardCode = N'" + EditText0.Value + "') "));
        }

        private void EditText0_KeyDownAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.CharPressed == 13)
            {
                if (Grid0.Rows.SelectedRows.Count == 0)
                {
                    return;
                }
                var x = Grid0.Rows.SelectedRows;
                var z = x.Item(0, BoOrderType.ot_RowOrder);
                //_landed._bpCardCode = Grid0.DataTable.GetValue("BP Code", z).ToString();
                //_landed._bpName = Grid0.DataTable.GetValue("BP Name", z).ToString();
                _landed.AfterChooseFromList(Grid0.DataTable.GetValue("BP Code", z).ToString(), Grid0.DataTable.GetValue("BP Name", z).ToString());

                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
            }
            else
            {
                Grid0.DataTable.ExecuteQuery(DiManager.QueryHanaTransalte("SELECT CardName as 'BP Name', CardCode as 'BP Code', Balance, VatIdUnCmp as 'Unified Federal Tax ID', GroupCode as 'Group Code', DebPayAcct as 'Account' FROM OCRD WHERE (CardName Like N'%" + EditText0.Value.Replace("'", "''") + "%') or (CardCode like N'%" + EditText0.Value.Replace("'", "''") + "%') or (VatIdUnCmp like N'%" + EditText0.Value.Replace("'", "''") + "%') or (DebPayAcct like N'%" + EditText0.Value.Replace("'", "''") + "%')"));
                Grid0.Rows.SelectedRows.Clear();
                Grid0.Rows.SelectedRows.Add(0);
            }
        }

        private void Grid0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            Grid0.Rows.SelectedRows.Clear();
            if (pVal.Row != -1)
            {
                Grid0.Rows.SelectedRows.Add(pVal.Row);
            }
        }
    }
}
