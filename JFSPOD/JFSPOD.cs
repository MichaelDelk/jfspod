/// <summary>
/// Class <c>JFSPOD</c> validates batches of JFSDOC documents for Kofax workflow.
/// 
/// This C# script replaces previous SBL version and removes references
/// to invoice date field that is no longer processed. Also removed 
/// save-and-skip logic.
/// 
/// IMPORTANT: odbcConnection string must be updated for whichever backend
/// connection is being used (i.e. production vs test environments).
/// 
/// 1. Enforces max length of index field Custno (field type J_ACCT#).
/// 2. Enforces max length of index field Invno (field type J_INV#).
/// 3. Validates Custno/Invno combination exists in backend table hhhordhp.
///    
/// Original: S4i Systems 2024-02-15 mtd
/// Modifications:
/// 
/// </summary>

using Kofax.AscentCapture.NetScripting;
using Kofax.Capture.CaptureModule.InteropServices;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Text;
using System.Data.Odbc;

namespace JFSPOD {
	
	
	[SuppressFieldEventsOnDocClose(false)]
	public class JFSPOD : DocumentValidationScript {
		
		[IndexFieldVariableAttribute("Custno")] 
		FieldScript Custno;
		
		[IndexFieldVariableAttribute("Invno")] 
		FieldScript Invno;

        #region S4i Work Fields
        // IMPORTANT: Change odbcConnection value as needed for desired environment.
        // string odbcConnectionString = "DSN=JFS;UID=edgar;PWD=edgar"; // Jordano production value
        // string odbcConnectionString = "DSN=SMRDEMO32;UID=kofaxdb;PWD=dbkofax"; // S4i Test value
        OdbcConnection DbConnection;
        string odbcConnectionString = "DSN=SMRDEMO32;UID=kofaxdb;PWD=dbkofax"; // S4i Test value

        // Capture index field values for use in database lookup.
        string keyCustno = " ";
        string keyInvno = " ";
        #endregion

        public JFSPOD(bool bIsValidation, string strUserID, string strLocaleName)
		: base(bIsValidation, strUserID, strLocaleName)
		{
            this.BatchLoading += JFSPOD_BatchLoading;
            this.BatchUnloading += JFSPOD_BatchUnloading;
            this.DocumentPreProcessing += JFSPOD_DocumentPreProcessing;
            this.DocumentPostProcessing += JFSPOD_DocumentPostProcessing;
		}

        #region Batch Load/Unload
        /// <summary>
        /// Called when a batch is closed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void JFSPOD_BatchUnloading(object sender, BatchEventArgs e)
        {
            Custno.FieldFormatting -= Custno_FieldFormatting;
            Custno.FieldPostProcessing -= Custno_FieldPostProcessing;
            Custno.FieldPreProcessing -= Custno_FieldPreProcessing;
            Invno.FieldFormatting -= Invno_FieldFormatting;
            Invno.FieldPostProcessing -= Invno_FieldPostProcessing;
            Invno.FieldPreProcessing -= Invno_FieldPreProcessing;
            DocumentPostProcessing -= JFSPOD_DocumentPostProcessing;
            DocumentPreProcessing -= JFSPOD_DocumentPreProcessing;
            BatchLoading -= JFSPOD_BatchLoading;
            BatchUnloading -= JFSPOD_BatchUnloading;

            // Close connection to backend database.
            this.DbConnection.Close();
        }
        /// <summary>
        /// Called when a batch is opened.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void JFSPOD_BatchLoading(object sender, BatchEventArgs e)
        {
            this.Custno.FieldPreProcessing += Custno_FieldPreProcessing;
            this.Custno.FieldPostProcessing += Custno_FieldPostProcessing;
            this.Custno.FieldFormatting += Custno_FieldFormatting;
            this.Invno.FieldPreProcessing += Invno_FieldPreProcessing;
            this.Invno.FieldPostProcessing += Invno_FieldPostProcessing;
            this.Invno.FieldFormatting += Invno_FieldFormatting;

            // Establish connection to backend database.
            this.DbConnection = new OdbcConnection(this.odbcConnectionString);
            this.DbConnection.Open();
        }
        #endregion

        #region Document Pre/Post Processing
        /// <summary>
        /// DocumentPreProcessing Called each time a new document is opened.
        /// To signal an error state, the Visual C# script can throw an exception during event handling. Three types of
        /// exceptions are available:
        ///		FatalErrorException
        ///		RejectAndSkipDocumentException
        ///		ValidationErrorException
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void JFSPOD_DocumentPreProcessing(object sender, PreDocumentEventArgs e)
        {
            // Initialize document work field values.
            this.keyCustno = " ";
            this.keyInvno = " ";
        }

        /// <summary>
        /// Handling document post index processing event.
        /// To signal an error state, the Visual C# script can throw an exception during event handling.
        /// Three types of exceptions are available:
        ///		FatalErrorException
        ///		RejectAndSkipDocumentException
        ///		ValidationErrorException
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void JFSPOD_DocumentPostProcessing(object sender, PostDocumentEventArgs e)
        {
        }
        #endregion

        #region Database Methods

        /// <summary>
        /// Count hhhordhp rows for requested customer/invoice combination.
        /// </summary>
        /// <param name="cusno"></param>
        /// <param name="invno"></param>
        private int Hhhordhp_Count(string custno, string invno)
        {
            string sqlCommandText;
            string tick = "'";
            int rowCount = 0;

            sqlCommandText = "SELECT COUNT(*) " 
                + "FROM zdemouser.hhhordhp "
                + "WHERE hhhinvn = " + tick + invno + tick
                + "  AND hhhcusn = " + tick + custno + tick;

            OdbcCommand DbCommand = this.DbConnection.CreateCommand();
            DbCommand.CommandText = sqlCommandText;
            OdbcDataReader DbReader = DbCommand.ExecuteReader();
            if (DbReader.Read())
            {
                rowCount = DbReader.GetInt32(0);
            }
            DbReader.Close();
            DbCommand.Dispose();

            return rowCount;
        }
        #endregion

        #region Validate Methods

        private void ValidateTextFieldLength(FieldScript oFieldScript)
        {
            string strValue = oFieldScript.IndexField.Value.Trim();

            if (strValue.Length > oFieldScript.IndexField.Length)
                throw new ValidationErrorException(string.Format("Character length exceeds maximum of {0}.", oFieldScript.IndexField.Length));
            oFieldScript.IndexField.Value = strValue;
        }
        private void ReportValidationMsg(string methodName)
        {
            string msgText = "methodName: " + methodName 
                + " | " + " keyCustno: " + this.keyCustno
                + " | " + " keyInvno: " + this.keyInvno;
            throw new ValidationErrorException(msgText);
        }
        private void ReportFatalMsg(string methodName)
        {
            string msgText = "methodName: " + methodName
                + " | " + " keyCustno: " + this.keyCustno
                + " | " + " keyInvno: " + this.keyInvno;
            throw new FatalErrorException(msgText);
        }
        #endregion

        #region Custno scripts
        private void Custno_FieldFormatting(object sender, FormatFieldEventArgs e)
        {
        }
        private void Custno_FieldPreProcessing(object sender, PreFieldEventArgs e)
        {
        }
        private void Custno_FieldPostProcessing(object sender, PostFieldEventArgs e)
        {
            // Enforce max length per field type on index field.
            ValidateTextFieldLength(sender as FieldScript);

            // Store document work field values.
            FieldScript fieldScript = sender as FieldScript;
            this.keyCustno = fieldScript.IndexField.Value.Trim();
        }
        #endregion

        #region Invno scripts
        private void Invno_FieldFormatting(object sender, FormatFieldEventArgs e)
        {
        }
        private void Invno_FieldPreProcessing(object sender, PreFieldEventArgs e)
        {
        }
        private void Invno_FieldPostProcessing(object sender, PostFieldEventArgs e)
        {

            // Enforce max length per field type on index field.
            ValidateTextFieldLength(sender as FieldScript);

            // Store document work field values.
            FieldScript fieldScript = sender as FieldScript;
            this.keyInvno = fieldScript.IndexField.Value.Trim();

            // Customer/Invoice must exist in HHHORDHP table.
            if (Hhhordhp_Count(this.keyCustno, this.keyInvno) < 1)
            {
                string msgText = "Customer  " + this.keyCustno
                    + " / Invoice " + this.keyInvno
                    + " not found in HHHORDHP table.";
                throw new ValidationErrorException(msgText);
            }
        }
        #endregion
    }
}
