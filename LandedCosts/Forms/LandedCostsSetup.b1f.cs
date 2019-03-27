using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace LandedCosts.Forms
{
    [FormAttribute("898", "Forms/LandedCostsSetup.b1f")]
    class LandedCostsSetup : SystemFormBase
    {
        public LandedCostsSetup()
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

        private SAPbouiCOM.Matrix Matrix0;

        public static Action<string, string> LandedCostsCode;


        private void Matrix0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Form landedCostSetup = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                Matrix landedCostSetupMatrix = (Matrix)landedCostSetup.Items.Item("3").Specific;
                var code = ((EditText)landedCostSetupMatrix.Columns.Item(3).Cells.Item(pVal.Row).Specific).Value;
                var name = ((EditText)landedCostSetupMatrix.Columns.Item(4).Cells.Item(pVal.Row).Specific).Value;
                LandedCostsCode.Invoke(code, name);
                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
            }
            catch (Exception e)
            {
 
            }
        }

        private void OnCustomInitialize()
        {

        }
    }
}
