using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace LandedCosts.Forms
{
    [FormAttribute("895", "Forms/TaxGroups.b1f")]
    class TaxGroups : SystemFormBase
    {
        public TaxGroups()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("3").Specific));
            this.Matrix0.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.Matrix0_DoubleClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

 

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Matrix Matrix0;
        public static Action<string> _taxCode;

        private void Matrix0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Form landedCostSetup = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                Matrix landedCostSetupMatrix = (Matrix)landedCostSetup.Items.Item("3").Specific;
                var code = ((EditText)landedCostSetupMatrix.Columns.Item(3).Cells.Item(pVal.Row).Specific).Value;
                _taxCode.Invoke(code);
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
            }
            catch (Exception)
            {

            }
        }
    }
}
