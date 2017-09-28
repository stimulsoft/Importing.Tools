using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;

namespace Import.Rdl
{
    public class StiWarning
    {
        #region Fields
        public string Message = null;
        public string Description = null;
        #endregion

        #region Static methods
        public static StiWarning GetWarningPicture(string objectName)
        {
            return new StiWarning(
                string.Format("Image '{0}': Impossible to get content", objectName),
                "When using CrystalReports from .Net, it is possible to get only some properties of report object." +
                " For example, it is possible to get only image size and location, it is impossible to get image content.");
        }

        public static StiWarning GetWarningDatabase(string objectName)
        {
            return new StiWarning(
                string.Format("Database '{0}': Could not identify the type", objectName),
                "When connecting to data bases CrystalReports often uses its own internal libraries." +
                " It is possible to get only some properties of connection from .Net and impossible to get ConnectionString." +
                " Therefore, not all data bases can be identified." +
                " By default, for not identified bases StiOleDbDatabase and ConnectionString is used without specifying provider.");
        }

        public static StiWarning GetWarningFormulaField(string objectName)
        {
            return new StiWarning(
                string.Format("FormulaField '{0}': Expression can be incorrect", objectName),
                "CrystalReports allows using expressions and formulas in FormulaFields." +
                " On the current moment parsing and syntax converting cannot be done, expressions is written 'as is'." +
                " Therefore, in many cases it is required further manual correction of expressions.");
        }
        #endregion

        #region this
        public StiWarning()
            : this(string.Empty, string.Empty)
        {
        }

        public StiWarning(string message, string description)
        {
            this.Message = message;
            this.Description = description;
        }
        #endregion
    }
}
