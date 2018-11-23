using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.SqlServer;
//using W32_ACEInterface-3.dll;


namespace BIS4._1_API_Application
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            initiateLogging();
            dp.OnNeedSomething              += dp_OnSomeNeeded;                 // delegate from dp
            dp.updateProgressBarsCurrent    += dp_updateProgressBarsCurrent;    // delegate from dp
            dp.updateProgressBarsMaximum    += dp_updateProgressBarsMaximum;    // delegate from dp
            this.lbl_APP_Result.Text = "*** Program / logging started ***";
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            loadJobList(TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JOBLIST.Text, true);
        }

#region INITIATE OBJECTS------------------------------------------------------------------------------

        API_RETURN_CODES_CS result;
        AccessEngine ace = new AccessEngine();
        DataProcessor dp = new DataProcessor();

        public virtual string   PB_ReadData     { get; set; }
        public virtual string   PB_ConvertData  { get; set; }
        public virtual string   PB_SendData     { get; set; }
        public string           dataType        = "";
        public string           msg             = "";
        public string           msgError        = "";
        public string           msgErrorType    = "";
        public string           jobOption       = "";
        public string           sendMethod      = "";
        public int              visibleFileName = 35;
        public int              visibleDgvRows  = 100;
        public int              count_PROCESSED = 0;
        public int              count_ADDED     = 0;
        public int              count_SKIPPED   = 0;
        public int              count_UPDATED   = 0;
        public int              count_DELETED   = 0;
        public int              count_ERRORS    = 0;
        public int              count_API_LOGIN_FAIL = 0;
        public bool             returnState     = false;
        public bool             F_API_ONLINE    = false;                                    // flag 'API is ONLINE'
        public bool             F_PROCESSED     = false;

        public DataTable    dtRequestDB         = new DataTable("dtRequestDB");
        public DataTable    dtInputFile         = new DataTable("dtInputFile");
        public DataTable    dtAssocFile         = new DataTable("dtAssocFile");
        public DataTable    dtOutputFile        = new DataTable("dtOutputFile");
        public DataTable    dtOutputData        = new DataTable("dtOutputData"); 
        public DataTable    dtAPISchemaFile     = new DataTable("dtAPISchema");
        public DataTable    dtReview            = new DataTable("dtReview");
        public DataTable    dtCompany           = new DataTable("dtCompany");
        public DataTable    dtAuthorization     = new DataTable("dtAuthorization");
        public DataTable    dtCard              = new DataTable("dtCard");
        public DataTable    dtJobList           = new DataTable("dtJobList");
        public DataTable    dtTmp               = new DataTable("dtTmp");
        public DataTable    dtInputFileTmp      = new DataTable("dtInputFileTmp");
        public DataTable    dtOutputFileTmp     = new DataTable("dtOutputFileTmp");
        public DataSet      aceData             = new DataSet("aceapi");
        
        private void FormMain_Load(object sender, EventArgs e)
        {
            controlsStatus_Blocked();
            aceData.Tables.Add(dtRequestDB);
            aceData.Tables.Add(dtInputFile);
            aceData.Tables.Add(dtAssocFile);
            aceData.Tables.Add(dtOutputFile);
            aceData.Tables.Add(dtAPISchemaFile);
            aceData.Tables.Add(dtJobList);

            string[] fields = null;
            string currentPath = "";
            
            
            fields = Application.StartupPath.ToString().Split((char)92);
            for (int i = 0; i < fields.Length - 1; i++) currentPath = currentPath + fields[i] + @"\";
            currentPath = currentPath.Substring(0, currentPath.Length - 1);
            TB_WORKFILESPATH.Text = currentPath;
            TB_CFG_FILE.Text = currentPath + @"\_config\" + "vtb.test.cfg.xml";

            Load_CFG_Files();

            string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JOBLIST.Text;
            TB_JLE_JOBLIST.Text = this.TB_JOBLIST.Text;
            loadJobList(jobListPath, true);
        }

#endregion------------------------------------------------------------------------------

#region SINGLE API OPERATIONS ---------------------------------------------------------------------------------

        public void btn_LOGIN_Click(object sender, EventArgs e)
        {
            result = ace.Login("BIS", "BIS", "BIS41RU");
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Login OK!");
        }
        public void btn_LOGOUT_Click(object sender, EventArgs e)
        {
            ace.Logout();
        }

        private void btn_COMP_ADD_Click(object sender, EventArgs e)
        {
            ACECompanies aceComp = new ACECompanies(ace);
            aceComp.COMPANYNO = textBoxC_NO.Text;               // index field, no duplicates allowed
            aceComp.NAME = textBoxC_NAME.Text;
            result = aceComp.Add();
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Added: " + result);
            else MessageBox.Show("NOT added!!" + result);
        }
        private void btn_COMP_UPDATE_Click(object sender, EventArgs e)
        {
            ACECompanies aceComp = new ACECompanies(ace);
            result = aceComp.Get(textBoxC_ID.Text);         // index field, no duplicates allowed
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                if (rb_READ.Checked) textBoxC_NO.Text = aceComp.COMPANYNO;
                if (rb_READ.Checked) textBoxC_NAME.Text = aceComp.NAME;
                // if (rb_READ.Checked) tb_AUTH_MACID.Text = aceAuth.MACID;
                MessageBox.Show("Exists: " + aceComp.NAME);
                aceComp.COMPANYNO = textBoxC_NO.Text;
                aceComp.NAME = textBoxC_NAME.Text;
                result = aceComp.Update();
                MessageBox.Show("Update: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }
        private void btn_COMP_DELETE_Click(object sender, EventArgs e)
        {
            ACECompanies aceComp = new ACECompanies(ace);
            result = aceComp.Get(textBoxC_ID.Text);
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                MessageBox.Show("Exists: " + aceComp.NAME);
                result = aceComp.Delete();
                MessageBox.Show("Delete: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }

        private void btn_AUTH_ADD_Click(object sender, EventArgs e)
        {
            ACEAuthorizations aceAuth = new ACEAuthorizations(ace);
            aceAuth.SHORTNAME = tb_AUTH_SN.Text;
            aceAuth.NAME = tb_AUTH_NM.Text;
            aceAuth.MACID = tb_AUTH_MACID.Text;
            // --- GET MACID BEFORE ADDING!
            var query = new ACEQuery(ace);
            query.Select("id", "devices", "DATEDELETED is NULL AND type=’MAC’");
            while (query.FetchNextRow())
            {
                ACEColumnValue MACID = null;
                // Get the MAC’s unique identifier
                query.GetRowData(0, MACID);
            }
            // --- OK, MACID is ready, go on
            result = aceAuth.Add();
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Added: " + result);
            else MessageBox.Show("NOT added!!" + result);
        }
        private void btn_AUTH_UPDATE_Click(object sender, EventArgs e)
        {
            ACEAuthorizations aceAuth = new ACEAuthorizations(ace);
            result = aceAuth.Get(tb_AUTH_ID.Text);            // index field, no duplicates allowed
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                if (rb_READ.Checked) tb_AUTH_SN.Text = aceAuth.SHORTNAME;
                if (rb_READ.Checked) tb_AUTH_NM.Text = aceAuth.NAME;
                if (rb_READ.Checked) tb_AUTH_MACID.Text = aceAuth.MACID;
                MessageBox.Show("Exists: " + aceAuth.SHORTNAME + "  " + aceAuth.NAME);
                aceAuth.SHORTNAME = tb_AUTH_SN.Text;
                aceAuth.NAME = tb_AUTH_NM.Text;
                aceAuth.MACID = tb_AUTH_MACID.Text;
                TB_ASIGNAUTH_AUTHID.Text = tb_AUTH_ID.Text;
                // --- GET MACID BEFORE UPDATING!
                var query = new ACEQuery(ace);
                query.Select("id", "devices", "DATEDELETED is NULL AND type=’MAC’");
                while (query.FetchNextRow())
                {
                    ACEColumnValue MACID = null;
                    // Get the MAC’s unique identifier
                    query.GetRowData(0, MACID);
                }
                // --- OK, MACID is ready, go on
                if (rb_UPDATE.Checked) result = aceAuth.Update();
                MessageBox.Show("Update: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }
        private void btn_AUTH_DELETE_Click(object sender, EventArgs e)
        {
            ACEAuthorizations aceAuth = new ACEAuthorizations(ace);
            result = aceAuth.Get(tb_AUTH_ID.Text);
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                MessageBox.Show("Exists: " + aceAuth.SHORTNAME + "  " + aceAuth.NAME);
                result = aceAuth.Delete();
                MessageBox.Show("Delete: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }

        private void btn_CARD_ADD_Click(object sender, EventArgs e)
        {
            ACECards aceCard = new ACECards(ace);
            aceCard.CARDNO = tb_CARDS_CARDNO.Text;
            aceCard.CODEDATA = tb_CARDS_CODEDATA.Text;
            result = aceCard.Add();
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Added: " + result);
            else MessageBox.Show("NOT added!!" + result);
        }
        private void btn_CARD_UPDATE_Click(object sender, EventArgs e)
        {
            ACECards aceCard = new ACECards(ace);

            result = aceCard.Get(tb_CARDS_CARDID.Text);            // index field, no duplicates allowed
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                if (rb_READ.Checked) tb_CARDS_CARDNO.Text = aceCard.CARDNO;
                if (rb_READ.Checked) tb_CARDS_CODEDATA.Text = aceCard.CODEDATA;
                if (rb_READ.Checked) TB_CARDS_PERSID.Text = aceCard.PERSID;
                MessageBox.Show("Exists: " + aceCard.CARDNO + "  " + aceCard.CODEDATA);
                aceCard.CARDNO = tb_CARDS_CARDNO.Text;
                aceCard.CODEDATA = tb_CARDS_CODEDATA.Text;
                aceCard.PERSID = TB_CARDS_PERSID.Text;
                TB_ASSIGNCARD_CARDID.Text = tb_CARDS_CARDID.Text;
                if (rb_UPDATE.Checked) result = aceCard.Update();
                MessageBox.Show("Update: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }
        private void btn_CARD_DELETE_Click(object sender, EventArgs e)
        {
            ACECards aceCard = new ACECards(ace);

            result = aceCard.Get(tb_CARDS_CARDID.Text);            // index field, no duplicates allowed
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                MessageBox.Show("Exists: " + aceCard.CARDNO + "  " + aceCard.CODEDATA);
                result = aceCard.Delete();
                MessageBox.Show("Delete: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }

        private void btn_PERS_ADD_Click(object sender, EventArgs e)
        {
            ACEPersons acePers = new ACEPersons(ace);
            acePers.FIRSTNAME = tb_PERS_FIRSTNAME.Text;
            acePers.LASTNAME = tb_PERS_LASTNAME.Text;
            acePers.COMPANYID = tb_PERS_COMPANYID.Text;
            acePers.PERSNO = tb_PERS_PERSNO.Text;
            int gender = 0;
            acePers.SEX = (ACESexT)gender;
            result = acePers.Add();
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Added: " + result);
            else MessageBox.Show("NOT added!!" + result);
        }
        private void btn_PERS_UPDATE_Click(object sender, EventArgs e)
        {
            ACEPersons acePers = new ACEPersons(ace);

            result = acePers.Get(tb_PERS_PERSID.Text);            // index field, no duplicates allowed
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                if (rb_READ.Checked) tb_PERS_FIRSTNAME.Text = acePers.FIRSTNAME;
                if (rb_READ.Checked) tb_PERS_LASTNAME.Text = acePers.LASTNAME;
                if (rb_READ.Checked) tb_PERS_COMPANYID.Text = acePers.COMPANYID;
                if (rb_READ.Checked) tb_PERS_PERSNO.Text = acePers.PERSNO;
                MessageBox.Show("Exists: " + acePers.FIRSTNAME + "  " + acePers.LASTNAME);
                acePers.FIRSTNAME = tb_PERS_FIRSTNAME.Text;
                acePers.LASTNAME = tb_PERS_LASTNAME.Text;
                acePers.COMPANYID = tb_PERS_COMPANYID.Text;
                acePers.PERSNO = tb_PERS_PERSNO.Text;
                TB_ASSIGNCARD_PERSID.Text = tb_PERS_PERSID.Text;
                TB_ASSIGNAUTH_PERSID.Text = tb_PERS_PERSID.Text;
                if (rb_UPDATE.Checked) result = acePers.Update();
                MessageBox.Show("Update: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }
        private void btn_PERS_DELETE_Click(object sender, EventArgs e)
        {
            ACEPersons acePers = new ACEPersons(ace);

            result = acePers.Get(tb_PERS_PERSID.Text);            // index field, no duplicates allowed
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                MessageBox.Show("Exists: " + acePers.FIRSTNAME + "  " + acePers.LASTNAME);
                result = acePers.Delete();
                MessageBox.Show("Delete: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }

        private void btn_VISITORS_ADD_Click(object sender, EventArgs e)
        {
            ACEVisitors aceVisitor = new ACEVisitors(ace);
            aceVisitor.FIRSTNAME = TB_VISITOR_FIRSTNAME.Text;
            aceVisitor.LASTNAME = TB_VISITOR_LASTNAME.Text;
            aceVisitor.COMPANYID = TB_VISITOR_COMPANYID.Text;
            int gender = 0;
            aceVisitor.SEX = (ACESexT)gender;
            result = aceVisitor.Add();
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Added: " + result);
            else MessageBox.Show("NOT added!!" + result);
        }
        private void btn_VISITORS_UPDATE_Click(object sender, EventArgs ee)
        {
            var aceVisitor = new ACEVisitors(ace);
            API_RETURN_CODES_CS result = aceVisitor.Get(TB_VISITOR_VISID.Text);
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                if (rb_READ.Checked) TB_VISITOR_FIRSTNAME.Text = aceVisitor.FIRSTNAME;
                if (rb_READ.Checked) TB_VISITOR_LASTNAME.Text = aceVisitor.LASTNAME;
                if (rb_READ.Checked) TB_VISITOR_COMPANYID.Text = aceVisitor.COMPANYID;
                if (rb_READ.Checked) TB_VISITOR_VISID.Text = aceVisitor.GetVisitorId();
                MessageBox.Show("Exists: " + aceVisitor.FIRSTNAME + "  " + aceVisitor.LASTNAME);
                aceVisitor.FIRSTNAME = TB_VISITOR_FIRSTNAME.Text;
                aceVisitor.LASTNAME = TB_VISITOR_LASTNAME.Text;
                aceVisitor.COMPANYID = TB_VISITOR_COMPANYID.Text;
                TB_ASSIGNCARD_VISID.Text = TB_VISITOR_VISID.Text;
                TB_ASSIGNAUTH_VISID.Text = TB_VISITOR_VISID.Text;
                if (rb_UPDATE.Checked) result = aceVisitor.Update();
                MessageBox.Show("Update: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }
        private void btn_VISITORS_DELETE_Click(object sender, EventArgs e)
        {
            ACEVisitors aceVisitor = new ACEVisitors(ace);

            result = aceVisitor.Get(tb_PERS_PERSID.Text);            // index field, no duplicates allowed
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
            {
                MessageBox.Show("Exists: " + aceVisitor.FIRSTNAME + "  " + aceVisitor.LASTNAME);
                result = aceVisitor.Delete();
                MessageBox.Show("Delete: " + result);
            }
            else MessageBox.Show("Absent: " + result);
        }

        private void btn_ASSIGN_CARD_PERSON_Click(object sender, EventArgs e)
        {
            ACEPersons acePers = new ACEPersons(ace);
            ACECards aceCard = new ACECards(ace);

            result = acePers.Get(TB_ASSIGNCARD_PERSID.Text);
            result = aceCard.Get(TB_ASSIGNCARD_CARDID.Text);
            result = acePers.AddCard(aceCard.GetCardId());

            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Success: " + result);
            else MessageBox.Show("Fail: " + result);
        }
        private void btn_ASSIGN_CARD_VISITOR_Click(object sender, EventArgs e)
        {
            ACEVisitors aceVisitor = new ACEVisitors(ace);
            ACECards aceCard = new ACECards(ace);

            result = aceVisitor.Get(TB_ASSIGNCARD_VISID.Text);
            result = aceCard.Get(TB_ASSIGNCARD_CARDID.Text);
            result = aceVisitor.AddCard(aceCard.GetCardId());

            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Success: " + result);
            else MessageBox.Show("Fail: " + result);
        }
        private void btn_ASSIGN_AUTH_PERSON_Click(object sender, EventArgs e)
        {
            ACEPersons acePers = new ACEPersons(ace);
            ACEAuthorizations aceAuth = new ACEAuthorizations(ace);
            ACEDateT aceDateFrom = new ACEDateT();
            ACEDateT aceDateTill = new ACEDateT();

            aceDateFrom.Set(Convert.ToUInt32(TB_ASSIGNAUTH_FROM.Text.Substring(6, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_FROM.Text.Substring(4, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_FROM.Text.Substring(0, 4)));
            aceDateTill.Set(Convert.ToUInt32(TB_ASSIGNAUTH_TO.Text.Substring(6, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_TO.Text.Substring(4, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_TO.Text.Substring(0, 4)));

            result = acePers.Get(TB_ASSIGNAUTH_PERSID.Text);
            result = aceAuth.Get(TB_ASIGNAUTH_AUTHID.Text);
            result = acePers.AddAuthorization(aceAuth.GetAuthorizationId(), aceDateFrom, aceDateTill);

            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Success: " + result);
            else MessageBox.Show("Fail: " + result);
        }
        private void btn_ASSIGN_AUTH_VISITOR_Click(object sender, EventArgs e)
        {
            ACEVisitors aceVisitor = new ACEVisitors(ace);
            ACEAuthorizations aceAuth = new ACEAuthorizations(ace);
            ACEDateT aceDateFrom = new ACEDateT();
            ACEDateT aceDateTill = new ACEDateT();

            aceDateFrom.Set(Convert.ToUInt32(TB_ASSIGNAUTH_FROM.Text.Substring(6, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_FROM.Text.Substring(4, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_FROM.Text.Substring(0, 4)));
            aceDateTill.Set(Convert.ToUInt32(TB_ASSIGNAUTH_TO.Text.Substring(6, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_TO.Text.Substring(4, 2)), Convert.ToUInt32(TB_ASSIGNAUTH_TO.Text.Substring(0, 4)));

            result = aceVisitor.Get(TB_ASSIGNAUTH_VISID.Text);
            result = aceAuth.Get(TB_ASIGNAUTH_AUTHID.Text);
            result = aceVisitor.AddAuthorization(aceAuth.GetAuthorizationId(), aceDateFrom, aceDateTill);

            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Success: " + result);
            else MessageBox.Show("Fail: " + result);
        }

        private void btn_TIMEMODEL_ADD_Click(object sender, EventArgs e)
        {
            ACETimeModels aceTimeModel = new ACETimeModels(ace);
            aceTimeModel.NAME = TB_TM_NAME.Text;
            aceTimeModel.DESCRIPTION = TB_TM_DESCRIPTION.Text;
            ACEDateT aceDate = new ACEDateT();
            //aceDate.Set((uint)07, (uint)03, (uint)1973);
            //aceDate.Set((uint)dayOfBirth.Day, (uint)dayOfBirth.Month, (uint)dayOfBirth.Year);
            aceDate.Set(Convert.ToUInt32(TB_TM_REFDATE.Text.Substring(6, 2)), Convert.ToUInt32(TB_TM_REFDATE.Text.Substring(4, 2)), Convert.ToUInt32(TB_TM_REFDATE.Text.Substring(0, 4)));
            aceTimeModel.REFDATE = aceDate;
            ACEBoolT flag = new ACEBoolT();
            flag.op_Assign(cb_IgnoreSpecialDays.Checked);
            aceTimeModel.IGNORESPECDAYS = flag;

            result = aceTimeModel.Add();
            if (API_RETURN_CODES_CS.API_SUCCESS_CS == result) MessageBox.Show("Added: " + result);
            else MessageBox.Show("NOT added!!" + result);
        }

#endregion------------------------------------------------------------------------------

#region SETUP FUNCTIONS------------------------------------------------------------------------------
        // login related
        private void btnLogin_Click(object sender, EventArgs e)
        {
            APILogin(TB_API_LOGIN_SERVER.Text, API_LOGIN_NAME.Text, API_LOGIN_PWD.Text);
        }
        public  bool APILogin(string api_server, string api_userID, string api_password)
        {
            API_RETURN_CODES_CS result = API_RETURN_CODES_CS.API_AUTHENTICAION_FAILED_CS;
            //if (lblLoginStatus.Text != "ONLINE")
            if (!F_API_ONLINE)
            {
                String server = api_server.Trim();
                String userId = api_userID.Trim();
                String password = api_password.Trim();
                Cursor.Current = Cursors.WaitCursor;

                result = ace.Login(userId, password, server);

                if (API_RETURN_CODES_CS.API_SUCCESS_CS == result)
                {
                    F_API_ONLINE = true;
                    this.lbl_APP_Result.Text = "API LOGIN SUCCESSFUL; " + result.ToString() + "; " + server + "; " + userId + "; " + password;
                    this.lblLoginStatus.ForeColor = System.Drawing.Color.Green;
                    this.lblLoginStatus2.ForeColor = System.Drawing.Color.Green;
                    this.lblLoginStatus.Text = "ONLINE";
                    this.btnAPILogin.Enabled = false;
                    this.TB_API_LOGIN_SERVER.ReadOnly = true;
                    this.API_LOGIN_NAME.ReadOnly = true;
                    this.API_LOGIN_PWD.ReadOnly = true;
                    Cursor.Current = Cursors.Default;
                    this.lblLoginStatus2.Text = this.lblLoginStatus.Text;
                    global.Data.logger(this.lbl_APP_Result.Text, "10000001");
                    return true;
                }
                else
                {
                    F_API_ONLINE = false;
                    this.lbl_APP_Result.Text = "API LOGIN FAILED;   " + server + "  " + userId + "  " + password + "  " + result.ToString();
                    this.lblLoginStatus.ForeColor = System.Drawing.Color.Red;
                    this.lblLoginStatus2.ForeColor = System.Drawing.Color.Red;
                    this.lblLoginStatus.Text = "OFFLINE";
                    btnAPILogin.Enabled = true;
                    this.TB_API_LOGIN_SERVER.ReadOnly = false;
                    this.API_LOGIN_NAME.ReadOnly = false;
                    this.API_LOGIN_PWD.ReadOnly = false;
                    Cursor.Current = Cursors.Default;
                    msgError = "API LOGIN FAILED! \r\n" +
                                "Check:  \r\n" +
                                "1. Login account  \r\n" +
                                "2. Server availability  \r\n" + 
                                "3. BIS Server is running";
                    this.lblLoginStatus2.Text = this.lblLoginStatus.Text;
                    global.Data.logger(this.lbl_APP_Result.Text, "10000001");
                    global.Data.logger(msgError, "10000001");
                    return false;
                }
            }
            return false;
        }
        private void btnLogout_Click(object sender, EventArgs e)
        {
            API_RETURN_CODES_CS result = API_RETURN_CODES_CS.API_AUTHENTICAION_FAILED_CS;
            //if (lblLoginStatus.Text == "ONLINE")
            if (F_API_ONLINE)
            {
                result = ace.Logout();
                this.lbl_APP_Result.Text = "API LOGOUT SUCCESSFUL; " + result.ToString();
                this.lblLoginStatus.ForeColor = System.Drawing.Color.Red;
                this.lblLoginStatus2.ForeColor = System.Drawing.Color.Red;
                this.lblLoginStatus.Text = "OFFLINE";
                btnAPILogin.Enabled = true;
                this.TB_API_LOGIN_SERVER.ReadOnly = false;
                this.API_LOGIN_NAME.ReadOnly = false;
                this.API_LOGIN_PWD.ReadOnly = false;
                Cursor.Current = Cursors.Default;
                this.lblLoginStatus2.Text = this.lblLoginStatus.Text;
                global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            }
        }
        private void lblLoginStatus_Click(object sender, EventArgs e)
        {
            this.lbl_APP_Result.Text = "Feature is under construction";
            MessageBox.Show("Last known Result code: ");
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            //if (lblLoginStatus.Text == "ONLINE")
            if (F_API_ONLINE)
            {
                API_RETURN_CODES_CS result = API_RETURN_CODES_CS.API_AUTHENTICAION_FAILED_CS;
                try
                {
                    result = ace.Logout();
                    this.lbl_APP_Result.Text = "API LOGOUT ON EXIT SUCCESSFUL; " + result.ToString();
                    global.Data.logger(this.lbl_APP_Result.Text, "10000001");
                }
                catch { }
            }
            base.OnClosing(e);
        }
        // sql connection related
        private void btn_GenerateConnection1String_Click(object sender, EventArgs e)
        {
            this.TB_DB_CONNECTION_1_STRING.Text =
            "Data Source=" + this.TB_CONNECTION_SERVER_1.Text + @"\" + this.TB_CONNECTION_SERVER1_INSTANCE.Text + ";"
            + "Initial Catalog=" + this.TB_CONNECTION_SERVER1_DB.Text + ";"
            + "Persist Security Info=True;"
            + "User ID=" + this.TB_CONNECTIONSERVER1_USERID.Text + ";"
            + "Password=" + this.TB_CONNECTION_SERVER1_PWD.Text;
            this.lbl_APP_Result.Text = "SQL CONNECTION #1 STRING is generated; " + this.TB_DB_CONNECTION_1_STRING.Text;
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
        }
        private void btn_GenerateConnection2String_Click(object sender, EventArgs e)
        {
            this.TB_DB_CONNECTION_2_STRING.Text =
            "Data Source=" + this.TB_CONNECTION_SERVER_2.Text + @"\" + this.TB_CONNECTION_SERVER2_INSTANCE.Text + ";"
            + "Initial Catalog=" + this.TB_CONNECTION_SERVER2_DB.Text + ";"
            + "Persist Security Info=True;"
            + "User ID=" + this.TB_CONNECTIONSERVER2_USERID.Text + ";"
            + "Password=" + this.TB_CONNECTION_SERVER2_PWD.Text;
            this.lbl_APP_Result.Text = "SQL CONNECTION #2 STRING is generated; " + this.TB_DB_CONNECTION_2_STRING.Text;
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
        }
        private void btn_SaveConnection1StringToFile_Click(object sender, EventArgs e)
        {
            if (dp.txtSaveToFile(TB_DB_CONNECTION_1_STRING.Text, TB_WORKFILESPATH.Text + @"\_dbConnections\" + TB_DBCONNFILE1.Text))
                msg = "CONNECTION STRING is saved to file; " + TB_DBCONNFILE1.Text;
            else
                msg = "CONNECTION STRING is NOT saved to file; " + TB_DBCONNFILE1.Text;
            this.lbl_APP_Result.Text = msg;
            global.Data.logger(msg, "10000001");
        }
        private void btn_SaveConnection2StringToFile_Click(object sender, EventArgs e)
        {
            if (dp.txtSaveToFile(TB_DB_CONNECTION_2_STRING.Text, TB_WORKFILESPATH.Text + @"\_dbConnections\" + TB_DBCONNFILE2.Text))
                msg = "CONNECTION STRING is saved to file; " + TB_DBCONNFILE2.Text;
            else
                msg = "CONNECTION STRING is NOT saved to file; " + TB_DBCONNFILE2.Text;
            this.lbl_APP_Result.Text = msg;
            global.Data.logger(msg, "10000001");
        }
        private void btn_TestConnString1_Click(object sender, EventArgs e)
        {
            string connFile = TB_WORKFILESPATH.Text + @"\_dbConnections\" + TB_DBCONNFILE1.Text;
            lbl_SQL_Connection1_Test_Result.ForeColor = System.Drawing.Color.Red;
            lbl_SQL_Connection1_Test_Result.Text = "";

            lbl_testRQ1.Text = checkDBConnection(connFile);
            if (!String.IsNullOrEmpty(lbl_testRQ1.Text))
            {
                lbl_SQL_Connection1_Test_Result.ForeColor = System.Drawing.Color.Green;
                lbl_SQL_Connection1_Test_Result.Text = DateTime.Now + ": Connection OK";
            }
            else
            {
                lbl_SQL_Connection1_Test_Result.ForeColor = System.Drawing.Color.Red;
                lbl_SQL_Connection1_Test_Result.Text = DateTime.Now + ": NO Connection";
            }
        }
        private void btn_TestConnString2_Click(object sender, EventArgs e)
        {
            string connFile = TB_WORKFILESPATH.Text + @"\_dbConnections\" + TB_DBCONNFILE2.Text;
            lbl_SQL_Connection2_Test_Result.ForeColor = System.Drawing.Color.Red;
            lbl_SQL_Connection2_Test_Result.Text = "";

            lbl_testRQ2.Text = checkDBConnection(connFile);
            if (!String.IsNullOrEmpty(lbl_testRQ2.Text))
            {
                lbl_SQL_Connection2_Test_Result.ForeColor = System.Drawing.Color.Green;
                lbl_SQL_Connection2_Test_Result.Text = DateTime.Now + ": Connection OK";
            }
            else
            {
                lbl_SQL_Connection2_Test_Result.ForeColor = System.Drawing.Color.Red;
                lbl_SQL_Connection2_Test_Result.Text = DateTime.Now + ": NO Connection";
            }
        }
        private void btn_TestSQLConnection_in_Click(object sender, EventArgs e)
        {
            string connFile = TB_DBCONNFILE_IN.Text;
            if (!String.IsNullOrEmpty(checkDBConnection(connFile))) btn_TestSQLConnection_in.BackColor = System.Drawing.Color.LawnGreen;
            else btn_TestSQLConnection_in.BackColor = System.Drawing.Color.Red;
        }
        private void btn_TestSQLConnection_conv_Click(object sender, EventArgs e)
        {
            string connFile = TB_DBCONNFILE_PROCESSOR.Text;
            if (!String.IsNullOrEmpty(checkDBConnection(connFile))) btn_TestSQLConnection_conv.BackColor = System.Drawing.Color.LawnGreen;
            else btn_TestSQLConnection_conv.BackColor = System.Drawing.Color.Red;
        }
        private void btn_TestSQLConnection_send_Click(object sender, EventArgs e)
        {
            string connFile = TB_DBCONNFILE_OUT.Text;
            if (!String.IsNullOrEmpty(checkDBConnection(connFile))) btn_TestSQLConnection_send.BackColor = System.Drawing.Color.LawnGreen;
            else btn_TestSQLConnection_send.BackColor = System.Drawing.Color.Red;
        }
        private string checkDBConnection(string connFile)
        {
            try
            {
                string sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\testConnection.sql";
                string sqlRq = dp.txtReadFromFile(sqlRqFile);
                string result = dp.dtFindDB(connFile, sqlRqFile, "", "", "", "");
                if (!String.IsNullOrEmpty(result))
                {
                    msg = "SQL DB is ACCESSIBLE:\r\n" + connFile + "\r\n";
                    global.Data.logger(msg, "10000001");
                    lbl_APP_Result.Text = msg;
                    return sqlRq + " " + result;
                }
                msg = "SQL DB is NOT ACCESSIBLE:\r\n" + connFile + "\r\n";
                global.Data.logger(msg, "10000001");
                lbl_APP_Result.Text = msg;
                return "";
            }
            catch
            {
                string msg = "SQL DB is NOT ACCESSIBLE:\r\n" + connFile + "\r\n";
                global.Data.logger(msg, "10000101");
                lbl_APP_Result.Text = msg;
                return "";
            }
        }
        // config and job files related
        public void btnLoadJoblist_Click(object sender, EventArgs e)
        {
            string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JOBLIST.Text;
            loadJobList(jobListPath, true);
        }
        public DataTable loadJobList(string jobListPath, bool Update)
        {
            DataTable dtJoblist = new DataTable("dtJoblist");
            dtJoblist.Rows.Clear(); dtJoblist.Clear(); dtJoblist.Columns.Clear();
            dtJoblist = dp.dtReadFromFile(jobListPath, true);
            dgv_JOBLIST.DataSource = dtJoblist;

            if (Update) try
                {
                    cbb_Job.DataSource = null;    
                    cbb_Job.Items.Clear();
                    for (int j = 0; j < dtJoblist.Rows.Count; j++)
                    {
                        cbb_Job.Items.Add(dtJoblist.Rows[j].ItemArray[0]);
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }

            this.TB_JOBSCOUNTER.Text = dtJoblist.Rows.Count.ToString();
            this.lbl_APP_Result.Text = "Joblist is loaded; Found: " + dtJoblist.Rows.Count.ToString() + " jobs";
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            return dtJoblist;
        }
        private void btn_LOAD_CFG_FILE_Click(object sender, EventArgs e)
        {
            ofdConfigFile.Title = "SET CONFIGURATION FILE";
            ofdConfigFile.Filter = "CONFIGURATION|*.cfg.xml";
            if (ofdConfigFile.ShowDialog() != DialogResult.OK) return;
            this.TB_CFG_FILE.Text = ofdConfigFile.FileName;
            this.lbl_APP_Result.Text = "CONFIGURATION FILE is set; " + ofdConfigFile.FileName;
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");

            Load_CFG_Files();

            string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JOBLIST.Text;
            this.TB_JLE_JOBLIST.Text = this.TB_JOBLIST.Text;
            loadJobList(jobListPath, true);
        }
        private void Load_CFG_Files()
        {
            // workfilespath (if empty in xml-file then curren folder is used from which application starts)
            if (dp.readXML(TB_CFG_FILE.Text, 0) != "")
                this.TB_WORKFILESPATH.Text = dp.readXML(TB_CFG_FILE.Text, 0);
            // API
            this.TB_LOG_API.Text = dp.readXML(TB_CFG_FILE.Text, 1);
            this.TB_LOGAPP.Text = dp.readXML(TB_CFG_FILE.Text, 2);
            this.TB_JOBLIST.Text = dp.readXML(TB_CFG_FILE.Text, 3);
            this.TB_API_LOGIN_SERVER.Text = dp.readXML(TB_CFG_FILE.Text, 4);
            this.API_LOGIN_NAME.Text = dp.readXML(TB_CFG_FILE.Text, 5);
            this.API_LOGIN_PWD.Text = dp.readXML(TB_CFG_FILE.Text, 6);
            // photos
            this.TB_PHOTOS_IN.Text = dp.readXML(TB_CFG_FILE.Text, 7);
            this.TB_PHOTOS_OUT.Text = dp.readXML(TB_CFG_FILE.Text, 8);
            // SQL
            this.TB_CONNECTION_SERVER_1.Text = dp.readXML(TB_CFG_FILE.Text, 9);
            this.TB_CONNECTION_SERVER1_INSTANCE.Text = dp.readXML(TB_CFG_FILE.Text, 10);
            this.TB_CONNECTION_SERVER1_DB.Text = dp.readXML(TB_CFG_FILE.Text, 11);
            this.TB_CONNECTIONSERVER1_USERID.Text = dp.readXML(TB_CFG_FILE.Text, 12);
            this.TB_CONNECTION_SERVER1_PWD.Text = dp.readXML(TB_CFG_FILE.Text, 13);
            this.TB_DBCONNFILE1.Text = dp.readXML(TB_CFG_FILE.Text, 14);
            this.TB_CONNECTION_SERVER_2.Text = dp.readXML(TB_CFG_FILE.Text, 15);
            this.TB_CONNECTION_SERVER2_INSTANCE.Text = dp.readXML(TB_CFG_FILE.Text, 16);
            this.TB_CONNECTION_SERVER2_DB.Text = dp.readXML(TB_CFG_FILE.Text, 17);
            this.TB_CONNECTIONSERVER2_USERID.Text = dp.readXML(TB_CFG_FILE.Text, 18);
            this.TB_CONNECTION_SERVER2_PWD.Text = dp.readXML(TB_CFG_FILE.Text, 19);
            this.TB_DBCONNFILE2.Text = dp.readXML(TB_CFG_FILE.Text, 20);
        }
        // redefine files manually
        private void btnSetDBConnFile_Click(object sender, EventArgs e)
        {
            ofdDBConnFile.Title = "SET DB CONNECTION FILE";
            ofdDBConnFile.Filter = "DB CONNECTION|*.dbc.csv";
            if (ofdDBConnFile.ShowDialog() != DialogResult.OK) return;
            this.TB_DBCONNFILE_IN.Text = ofdDBConnFile.FileName;
            this.lbl_APP_Result.Text = "DB CONNECTION FILE is set; " + ofdDBConnFile.FileName;
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
        }
        private void btnSetDBRqFile_Click(object sender, EventArgs e)
        {
            ofdDBRqFile.Title = "SET SQL-SCRIPT FILE";
            ofdDBRqFile.Filter = "SQL-SCRIPT|*.sql";
            if (ofdDBRqFile.ShowDialog() != DialogResult.OK) return;
            this.TB_SQLFILE_IN.Text = ofdDBRqFile.FileName;
            this.lbl_APP_Result.Text = "SQL-SCRIPT FILE is set; " + ofdDBRqFile.FileName;
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
        }
        private void btnSetInputFile_Click(object sender, EventArgs e)
        {
            this.TB_INPUTFILE.Text = "";
            ofdInputFile.Title = "SET INPUT FILE";
            ofdInputFile.Filter = "INPUT|*.in.csv";
            if (ofdInputFile.ShowDialog() != DialogResult.OK) return;
            this.TB_INPUTFILE.Text = ofdInputFile.FileName;
            this.lbl_APP_Result.Text = "INPUT FILE is set: ..." + ofdInputFile.FileName.Substring(ofdInputFile.FileName.Length - visibleFileName, visibleFileName);
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            dgvInputFile.DataSource = null;
            //ReadInputFile(this.TB_INPUTFILE.Text, cb_HasHeaders.Checked);
        }
        private void btnSetAPIAssociationsFile_Click(object sender, EventArgs e)
        {
            this.TB_APIASSOCFILE.Text = "";
            ofdOpenAPIAssocFile.Title = "SET API ASSOCIATION FILE";
            ofdOpenAPIAssocFile.Filter = "API ASSOCIATION|*.aaf.csv";
            if (ofdOpenAPIAssocFile.ShowDialog() != DialogResult.OK) return;
            this.TB_APIASSOCFILE.Text = ofdOpenAPIAssocFile.FileName;
            this.lbl_APP_Result.Text = "API ASSOCIATION FILE is set: ..." + ofdOpenAPIAssocFile.FileName.Substring(ofdOpenAPIAssocFile.FileName.Length - visibleFileName, visibleFileName);
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            dgvAssocFile.DataSource = null;
            //ReadAssocFile(this.TB_APIASSOCFILE.Text);
        }
        private void btnSetOutputFile_Click(object sender, EventArgs e)
        {
            this.TB_OUTPUTFILE.Text = "";
            ofdOutputFile.Title = "SET OUTPUT FILE";
            ofdOutputFile.Filter = "OUTPUT|*.out.csv";
            if (ofdOutputFile.ShowDialog() != DialogResult.OK) return;
            this.TB_OUTPUTFILE.Text = ofdOutputFile.FileName;
            this.lbl_APP_Result.Text = "OUTPUT FILE is set: ..." + ofdOutputFile.FileName.Substring(ofdOutputFile.FileName.Length - visibleFileName, visibleFileName);
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            dgvOutputFile.DataSource = null;
            //ReadOutputFile(this.TB_OUTPUTFILE.Text);
        }
        private void btnSetAPISchemaFile_Click(object sender, EventArgs e)
        {
            this.TB_APISCHEMAFILE.Text = "";
            ofdAPISchemaFile.Title = "SET API SCHEMA FILE";
            ofdAPISchemaFile.Filter = "API SCHEMA|*.asf.csv";
            if (ofdAPISchemaFile.ShowDialog() != DialogResult.OK) return;
            this.TB_APISCHEMAFILE.Text = ofdAPISchemaFile.FileName;
            this.lbl_APP_Result.Text = "API SCHEMA FILE is set: ..." + ofdAPISchemaFile.FileName.Substring(ofdAPISchemaFile.FileName.Length - visibleFileName, visibleFileName);
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            dgvAPISchemaFile.DataSource = null;
            //ReadApiSchemaFile(this.TB_APISCHEMAFILE.Text);
        }
        // logs related
        private void initiateLogging()
        {
            dp.RAM_LOGLength = Convert.ToInt32(TB_RAMLOGLength.Text);
            dp.LOG_Overwrite = cb_LogOverwrite.Checked;
            
            dp.LOG_Allowed = "";
            dp.LOG_Allowed = dp.LOG_Allowed + "1";                                                                          // Console logging is always enabled
            if (cb_LogRAM.Checked) dp.LOG_Allowed = dp.LOG_Allowed + "1"; else dp.LOG_Allowed = dp.LOG_Allowed + "0";       // RAM logging
            dp.LOG_Allowed = dp.LOG_Allowed + "0";                                                                          // Reserved
            dp.LOG_Allowed = dp.LOG_Allowed + "0";                                                                          // Reserved
            if (cb_LogEXT.Checked) dp.LOG_Allowed = dp.LOG_Allowed + "1"; else dp.LOG_Allowed = dp.LOG_Allowed + "0";       // Extended Application logging (data processor messages as well)
            if (cb_LogSRV.Checked) dp.LOG_Allowed = dp.LOG_Allowed + "1"; else dp.LOG_Allowed = dp.LOG_Allowed + "0";       // Service logging
            if (cb_LogAPI.Checked) dp.LOG_Allowed = dp.LOG_Allowed + "1"; else dp.LOG_Allowed = dp.LOG_Allowed + "0";       // API logging
            if (cb_LogAPP.Checked) dp.LOG_Allowed = dp.LOG_Allowed + "1"; else dp.LOG_Allowed = dp.LOG_Allowed + "0";       // Application logging

            dp.logger("SETUP; LOG SETTINGS; FLAGS: " + dp.LOG_Allowed + "; RAMLog max: " + TB_RAMLOGLength.Text, "11000001");
        }

#endregion -----------------------------------------------------------------------------------------

#region GET DATA AND FILE OPERATIONS---------------------------------------------------------------------------------

        private void btnRequestDB_Click(object sender, EventArgs e)
        {
            RequestDB(this.TB_DBCONNFILE_IN.Text, this.TB_SQLFILE_IN.Text);
        }
        public DataTable RequestDB(string DBCONNFILEOPEN, string SQLRQFILE)
        {
            this.TB_INPUTFILE_ROWS.Text = ""; this.TB_INPUTFILE_COLUMNS.Text = "";
            dtInputFile.Clear(); dtInputFile.Columns.Clear();
            dgvInputFile.DataSource = null;
            if (File.Exists(DBCONNFILEOPEN) && File.Exists(SQLRQFILE))
            {
                dtInputFile = dp.dtReadDB(DBCONNFILEOPEN, SQLRQFILE, "", "", "", "");
                if (cb_DisplayData_Input.Checked) dgvInputFile.DataSource = dtInputFile;
                this.lbl_APP_Result.Text = "INPUT DATA is read from DB; ..." + DBCONNFILEOPEN.Substring(DBCONNFILEOPEN.Length - visibleFileName, visibleFileName);
                this.TB_INPUTFILE_ROWS.Text = (dtInputFile.Rows.Count + 1).ToString();
                this.TB_INPUTFILE_COLUMNS.Text = dtInputFile.Columns.Count.ToString();
                global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            }
            else
            {
                global.Data.logger("RequestDB; Connection or request file is wrong or not found", "10000100");
            }
            return dtRequestDB;
        }
        private void btnSaveInputFile_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(this.TB_INPUTFILE.Text))
            {
                dp.dtSaveToFile(dtInputFile, this.TB_INPUTFILE.Text, true);
                //ReadInputFile(this.TB_INPUTFILE.Text, cb_HasHeaders.Checked);
                this.lbl_APP_Result.Text = "INPUT FILE is saved; ..." + TB_INPUTFILE.Text.Substring(TB_INPUTFILE.Text.Length - visibleFileName, visibleFileName);
                global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            }
        }
        private void btnReadInputFile_Click(object sender, EventArgs e)
        {
            ReadInputFile(this.TB_INPUTFILE.Text, cb_HasHeaders.Checked);
        }
        public DataTable ReadInputFile(string INPUTFILE, bool F_HASHEADERS)
        {
            this.TB_INPUTFILE_ROWS.Text = ""; this.TB_INPUTFILE_COLUMNS.Text = "";
            dtInputFile.Clear(); dtInputFile.Columns.Clear();
            dtInputFileTmp.Clear(); dtInputFileTmp.Columns.Clear();
            dgvInputFile.DataSource = null;
            if (File.Exists(INPUTFILE))
            {
                dtInputFile                     = dp.dtReadFromFile(INPUTFILE, F_HASHEADERS);
                if (cb_DisplayData_Input.Checked)
                {
                    if (dtInputFile.Rows.Count <= visibleDgvRows) dgvInputFile.DataSource = dtInputFile;
                    else
                    {
                        dtInputFileTmp = dtInputFile.Clone();
                        for (int i = 0; i < visibleDgvRows; i++) dtInputFileTmp.ImportRow(dtInputFile.Rows[i]);
                        dgvInputFile.DataSource = dtInputFileTmp;
                    }
                }
                this.lbl_APP_Result.Text        = "INPUT FILE is read; ..." + TB_INPUTFILE.Text.Substring(TB_INPUTFILE.Text.Length - visibleFileName, visibleFileName);
                this.TB_INPUTFILE_ROWS.Text     = (dtInputFile.Rows.Count + 1).ToString();
                this.TB_INPUTFILE_COLUMNS.Text  = dtInputFile.Columns.Count.ToString();
                global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            }
            return dtInputFile;
        }
        private void btnReadAssociationsFile_Click(object sender, EventArgs e)
        {
            ReadAssocFile(this.TB_APIASSOCFILE.Text);
        }
        public DataTable ReadAssocFile(string APIASSOCFILE)
        {
            this.TB_ASSOCFILE_ROWS.Text = ""; this.TB_ASSOCFILE_COLUMNS.Text = "";
            dtAssocFile.Clear(); dtAssocFile.Columns.Clear();
            dgvAssocFile.DataSource = null;
            if (File.Exists(APIASSOCFILE))
            {
                dtAssocFile                     = dp.dtReadFromFile(APIASSOCFILE, true);
                dgvAssocFile.DataSource         = dtAssocFile;
                this.lbl_APP_Result.Text        = "ASSOCIATIONS FILE is read; ..." + APIASSOCFILE.Substring(APIASSOCFILE.Length - visibleFileName, visibleFileName);
                this.TB_ASSOCFILE_ROWS.Text     = (dtAssocFile.Rows.Count + 1).ToString();
                this.TB_ASSOCFILE_COLUMNS.Text  = dtAssocFile.Columns.Count.ToString();
                global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            }
            return dtAssocFile;
        }
        private void btnUpdateAssocFile_Click(object sender, EventArgs e)
        {
            dp.dgvSaveToFile(dgvAssocFile, this.TB_APIASSOCFILE.Text);
            this.lbl_APP_Result.Text = "ASSOCIATIONS FILE is updated; ..." + TB_APIASSOCFILE.Text.Substring(TB_APIASSOCFILE.Text.Length - visibleFileName, visibleFileName);
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
        }
        private void btnReadOutputFile_Click(object sender, EventArgs e)
        {
            ReadOutputFile(this.TB_OUTPUTFILE.Text);
        }
        public DataTable ReadOutputFile(string OUTPUTFILE)
        {
            dataType = "READ OUTPUT FILE";
            this.TB_OUTPUTFILE_ROWS.Text = ""; this.TB_OUTPUTFILE_COLUMNS.Text = "";
             dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
             dtOutputFileTmp.Clear(); dtOutputFileTmp.Columns.Clear();  
                if (true)
                    try
                    {
                        dgvOutputFile.DataSource = null;
                        if (File.Exists(OUTPUTFILE))
                        {
                            dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);
                            if (cb_DisplayData_Output.Checked)
                            {
                                if (dtOutputFile.Rows.Count <= visibleDgvRows) dgvOutputFile.DataSource = dtOutputFile;
                                else
                                {
                                    dtOutputFileTmp = dtOutputFile.Clone();
                                    for (int i = 0; i < visibleDgvRows; i++) dtOutputFileTmp.ImportRow(dtOutputFile.Rows[i]);
                                    dgvOutputFile.DataSource = dtOutputFileTmp;
                                }
                            }
                            this.lbl_APP_Result.Text = "OUTPUT FILE is read; ..." + OUTPUTFILE.Substring(OUTPUTFILE.Length - visibleFileName, visibleFileName);
                            this.TB_OUTPUTFILE_ROWS.Text = (dtOutputFile.Rows.Count + 1).ToString();
                            this.TB_OUTPUTFILE_COLUMNS.Text = (dtOutputFile.Columns.Count).ToString();
                            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
                        }
                        else
                        {
                            global.Data.logger("ReadOutputFile. File does not exist: " + OUTPUTFILE, "10000100");
                        }
                    }
                    catch (Exception e)
                    {
                        msgError = "ReadOutputFile. \r\n" +
                                    dataType + "\r\n" +
                                    e.Message + "\r\n" +
                                    "1. Check if file exists, accessible and has proper format\r\n" +
                        "Options: Abort the application or Retry to start new operation or Ignore this message";
                        msgErrorType = "GENERAL EXCEPTION";
                        global.Data.logger(msgError, "10000100");
                        DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                        if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                        if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); return dtOutputFile; }
                        if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); return dtOutputFile; }
                    }
            return dtOutputFile;
        }
        private void btnUpdateOutputFile_Click(object sender, EventArgs e)
        {
            dp.dgvSaveToFile(dgvOutputFile, this.TB_OUTPUTFILE.Text);
            //ReadOutputFile(this.TB_OUTPUTFILE.Text);
            this.lbl_APP_Result.Text = "OUTPUT FILE is updated; ..." + TB_OUTPUTFILE.Text.Substring(TB_OUTPUTFILE.Text.Length - visibleFileName, visibleFileName);
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
        }
        private void btnReadAPISchemaFile_Click(object sender, EventArgs e)
        {
            ReadApiSchemaFile(this.TB_APISCHEMAFILE.Text);
        }
        public DataTable ReadApiSchemaFile(string APISCHEMAFILE)
        {
            this.TB_APISCHEMAFILE_ROWS.Text = ""; this.TB_APISCHEMAFILE_COLUMNS.Text = "";
            dtAPISchemaFile.Clear(); dtAPISchemaFile.Columns.Clear();
            dgvAPISchemaFile.DataSource = null;
            if (File.Exists(APISCHEMAFILE))
            {
                dtAPISchemaFile                     = dp.dtReadFromFile(APISCHEMAFILE, true);
                dgvAPISchemaFile.DataSource         = dtAPISchemaFile;
                this.lbl_APP_Result.Text            = "API SCHEMA FILE is read; ..." + APISCHEMAFILE.Substring(APISCHEMAFILE.Length - visibleFileName, visibleFileName);
                this.TB_APISCHEMAFILE_ROWS.Text     = (dtAPISchemaFile.Rows.Count + 1).ToString();
                this.TB_APISCHEMAFILE_COLUMNS.Text  = (dtAPISchemaFile.Columns.Count).ToString();
                global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            }
            return dtAPISchemaFile;
        }
        private void btnUpdateAPISchemaFile_Click(object sender, EventArgs e)
        {
            dp.dgvSaveToFile(dgvAPISchemaFile, this.TB_APISCHEMAFILE.Text);
            //ReadApiSchemaFile(this.TB_APISCHEMAFILE.Text);
            this.lbl_APP_Result.Text = "API SCHEMA FILE is updated; ..." + TB_APISCHEMAFILE.Text.Substring(TB_APISCHEMAFILE.Text.Length - visibleFileName, visibleFileName);
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
        }

#endregion ------------------------------------------------------------------------------

#region CONVERSION OPERATIONS---------------------------------------------------------------------------------

        private void btnCONVERT_Click(object sender, EventArgs e)
        {
            CONVERT_Start();
        }
        private void CONVERT_Start()
        {
            int INPUTFILE_ROWS = 0;
            int APIASSOCFILE_ROWS = 0;
            if (true) try
                {
                    if (!String.IsNullOrEmpty(TB_INPUTFILE.Text)) INPUTFILE_ROWS = System.IO.File.ReadAllLines(TB_INPUTFILE.Text).Length;
                    if (!String.IsNullOrEmpty(TB_APIASSOCFILE.Text)) APIASSOCFILE_ROWS = System.IO.File.ReadAllLines(TB_APIASSOCFILE.Text).Length;
                }
                catch { MessageBox.Show("WRONG INPUT\n1. Check if input file exists", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

            if (INPUTFILE_ROWS > 0 && APIASSOCFILE_ROWS > 0)
                //try
                {
                    TB_INPUTFILE_ROWS.Text = INPUTFILE_ROWS.ToString();
                    TB_ASSOCFILE_ROWS.Text = APIASSOCFILE_ROWS.ToString();
                    //dgvOutputFile.DataSource = null;
                    dtOutputFile = CONVERT(this.TB_INPUTFILE.Text, this.TB_OUTPUTFILE.Text, this.TB_APIASSOCFILE.Text);
                    if (cb_FILE_PROCESSOR.Checked && !String.IsNullOrEmpty(TB_PROCESSOR.Text))
                    {
                        PROCESS(this.TB_PROCESSOR.Text, this.TB_OUTPUTFILE.Text, this.TB_AUX_FILE.Text, this.TB_DBCONNFILE_PROCESSOR.Text);
                        this.lbl_APP_Result.Text = "CONVERT (phase 4); OUTPUT FILE IS SAVED; Conversion completed";
                        global.Data.logger(this.lbl_APP_Result.Text, "10000001");
                    }
                    //ReadOutputFile(this.TB_OUTPUTFILE.Text);
                }
                //catch { MessageBox.Show("WRONG INPUT\n1. Check if output and schema files exist;\n2. Check if start line and number of lines are defined properly", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

        }
        public DataTable CONVERT(string INPUTFILE, string OUTPUTFILE, string ASSOCFILE)
        {
            DataTable dt = new DataTable("dt");
            dt.Clear(); dt.Columns.Clear();

            if (!String.IsNullOrEmpty(INPUTFILE) && !String.IsNullOrEmpty(ASSOCFILE))
                try
                {
                    dtInputFile = dp.dtReadFromFile(INPUTFILE, cb_HasHeaders.Checked);
                    dtAssocFile = dp.dtReadFromFile(ASSOCFILE, true);

                    //int i = 0; 
                    int j = 0; int k = 1; int l = 0; int m = 0; // i = column, j = row, k = rows counter in assoc file, l = outputColumnfromAssoc
                    dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
                    dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtInputFile.Rows.Count, -1);

                    // create headers
                    for (k = 0; k < dtAssocFile.Rows.Count; k++) dt.Columns.Add(dtAssocFile.Rows[k].ItemArray[2].ToString());

                    // transfer data            
                    for (j = 0; j < dtInputFile.Rows.Count; j++)
                    {
                        global.Data.logger("CONVERT (phase 1). Processing row #" + j.ToString() + "----------------------", "10001000"); // show # of row in console
                        l = 0; m = 0;
                        dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);

                        DataRow newRow = null;
                        newRow = dt.NewRow();

                        for (k = 0; k < dtAssocFile.Rows.Count; k++)
                        {
                            l = Convert.ToInt32(dtAssocFile.Rows[k].ItemArray[1].ToString());
                            if (dtAssocFile.Rows[k].ItemArray[0].ToString() != "")
                            {
                                m = Convert.ToInt32(dtAssocFile.Rows[k].ItemArray[0].ToString());
                                newRow[l] = dtInputFile.Rows[j].ItemArray[m].ToString();
                            }
                            else
                            {
                                newRow[l] = "";
                                string connString, sqlRq, Data1, Data2;
                                connString = sqlRq = Data1 = Data2 = "";

                                if (dtAssocFile.Rows[k].ItemArray[4].ToString() != "")                                  // Request of external data
                                {
                                    if (dtAssocFile.Rows[k].ItemArray[5].ToString() != "")
                                    {
                                        int extColumn = Convert.ToInt32(dtAssocFile.Rows[k].ItemArray[5].ToString());  // This column from input file contains data for external request
                                        Data1 = dtInputFile.Rows[j].ItemArray[extColumn].ToString();
                                    }
                                    if (dtAssocFile.Rows[k].ItemArray[6].ToString() != "")
                                    {
                                        int extColumn = Convert.ToInt32(dtAssocFile.Rows[k].ItemArray[6].ToString());  // This column from input file contains data for external request
                                        Data2 = dtInputFile.Rows[j].ItemArray[extColumn].ToString();
                                    }

                                    string connFile = TB_DBCONNFILE_PROCESSOR.Text;
                                    string sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + dtAssocFile.Rows[k].ItemArray[4].ToString();

                                    newRow[l] = dp.dtFindDB(connFile, sqlRqFile, Data1, Data2, "", "");
                                    global.Data.logger(
                                        "CONVERT (phase 1). dtFindDB----- \r\n" +
                                        "connFile: ..."    + connFile.Substring(connFile.Length    - visibleFileName, visibleFileName) + "\r\n" +
                                        "sqlRqFile: ..."   + sqlRqFile.Substring(sqlRqFile.Length  - visibleFileName, visibleFileName) + "\r\n" +
                                        "Data1: "       + Data1 + "\r\n" +
                                        "Data2: "       + Data2 + "\r\n" +
                                        "found: "       + newRow[l] + "\r\n"
                                        ,"10001000");
                                }
                            }
                        }
                        dt.Rows.Add(newRow);
                    }
                }
                catch (Exception e) { MessageBox.Show("CONVERT (phase 1); \nConversion failed; \n" + e.Message, "GENERAL EXCEPTION", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

            this.lbl_APP_Result.Text = "CONVERT (phase 1); CONVERSION COMPLETED; New output file will be created";
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");

            // save datatable to file
            dp.dtSaveToFile(dt, OUTPUTFILE, true);
            this.lbl_APP_Result.Text = "CONVERT (phase 2); OUTPUT FILE IS SAVED; File processor will be engaged if needed";
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            return dt;
        }
        private void PROCESS(string PROCESSOR, string OUTPUTFILE, string AUX_FILE, string DBCONNFILE)
        {
            this.lbl_APP_Result.Text = "CONVERT (phase 3); FILE PROCESSOR IS ENGAGED; " + PROCESSOR;
            global.Data.logger(this.lbl_APP_Result.Text, "10000001");
            switch (PROCESSOR)
            {
                case "AUTHORIZATIONS_CLONE_IF_MAC_DIFFERS":
                    AUTHORIZATIONS_CLONE_IF_MAC_DIFFERS(PROCESSOR, OUTPUTFILE, DBCONNFILE);
                    break;
                case "AUTHORIZATIONS_FILL_BIS41_AUTHID":
                    AUTHORIZATIONS_FILL_BIS41_AUTHID(PROCESSOR, OUTPUTFILE, DBCONNFILE);
                    break;
                //case "FIND_BIS41_PERSID":
                    //FIND_BIS41_PERSID(PROCESSOR, OUTPUTFILE, DBCONNFILE);
                    //break;
                case "FILL_CARD_ATTRIBUTES":
                    FILL_CARD_ATTRIBUTES(PROCESSOR, OUTPUTFILE, DBCONNFILE);
                    break;
                case "FIND_BIS41_AUTHID":
                    FIND_BIS41_AUTHID(PROCESSOR, OUTPUTFILE, DBCONNFILE);
                    break;
                case "FIND_BIS41_COMPANYID_EMPLOYEES":
                    FIND_BIS41_COMPANYID_EMPLOYEES(PROCESSOR, OUTPUTFILE, AUX_FILE, DBCONNFILE);
                    break;
                case "FIND_BIS41_COMPANYID_VISITORS":
                    FIND_BIS41_COMPANYID_VISITORS(PROCESSOR, OUTPUTFILE, AUX_FILE, DBCONNFILE);
                    break;
                default:
                    break;
            }
        }
        private void AUTHORIZATIONS_CLONE_IF_MAC_DIFFERS(string PROCESSOR, string OUTPUTFILE, string DBCONNFILE)
        {
            DataTable dtOutputFile = new DataTable("dtOutputFile");
            dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
            DataTable dtOutputData = new DataTable("dtOutputData");
            dtOutputData.Clear(); dtOutputData.Columns.Clear();

            dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
            dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
            string connFile = TB_DBCONNFILE_PROCESSOR.Text;
 
            dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);

            dtOutputData.Columns.Add("23_AUTHID");
            dtOutputData.Columns.Add("23_SHORTNAME");
            dtOutputData.Columns.Add("23_NAME");
            dtOutputData.Columns.Add("41_TMID");
            dtOutputData.Columns.Add("CLIENTID");
            dtOutputData.Columns.Add("SPECIALFUNCTIONID");
            dtOutputData.Columns.Add("41_MAC_ID");
            dtOutputData.Columns.Add("41_SHORTNAME");
            dgvOutputFile.DataSource = null;

            for (int j = 0; j < dtOutputFile.Rows.Count; j++)
            {
                dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
                global.Data.logger("CONVERT (phase 3). Processing row #" + j.ToString() + "-----------------", "10001000");    // show # of row in console

                DataRow dr = null;
                dr = dtOutputData.NewRow();
                dr[0] = dtOutputFile.Rows[j].ItemArray[0].ToString().Trim();
                dr[1] = dtOutputFile.Rows[j].ItemArray[1].ToString().Trim();
                dr[2] = dtOutputFile.Rows[j].ItemArray[2].ToString().Trim();
                dr[3] = dtOutputFile.Rows[j].ItemArray[3].ToString().Trim();
                dr[4] = dtOutputFile.Rows[j].ItemArray[4].ToString().Trim();
                dr[5] = dtOutputFile.Rows[j].ItemArray[5].ToString().Trim();
                dr[6] = dtOutputFile.Rows[j].ItemArray[6].ToString().Trim();
                dr[7] = dtOutputFile.Rows[j].ItemArray[1].ToString().Trim() + @"   @" + dtOutputFile.Rows[j].ItemArray[11].ToString().Trim();

                bool F_RECORD_EXISTS = false;
                for (int i = 0; i < dtOutputData.Rows.Count; i++)
                {
                    if (dr[7].ToString() == dtOutputData.Rows[i].ItemArray[7].ToString()) F_RECORD_EXISTS = true;
                }
                if (!F_RECORD_EXISTS) dtOutputData.Rows.Add(dr);
            }
            dp.dtSaveToFile(dtOutputData, OUTPUTFILE, true);
            dtOutputFile.Dispose();
            dtOutputData.Dispose();
        }
        private void AUTHORIZATIONS_FILL_BIS41_AUTHID(string PROCESSOR, string OUTPUTFILE, string DBCONNFILE)
        {
            DataTable dtOutputFile = new DataTable("dtOutputFile");
            dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
            DataTable dtOutputData = new DataTable("dtOutputData");
            dtOutputData.Clear(); dtOutputData.Columns.Clear();
            DataTable dtAuth41Data = new DataTable("dtAuth41Data");
            dtAuth41Data.Clear(); dtAuth41Data.Columns.Clear();

            dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
            dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
            string connFile = TB_DBCONNFILE_PROCESSOR.Text;
            string sqlRqFile = "";

            dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);

            sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.AUTHORIZATIONS_FILL_BIS41_AUTHID.sql";
            dtAuth41Data = dp.dtReadDB(DBCONNFILE, sqlRqFile, "", "", "", ""); // BIS4.1 AUTHID col0, SHORTNAME col2
            dgvDisplayOutput.DataSource = null;

            dtOutputData.Columns.Add("23_SHORTNAME");
            dtOutputData.Columns.Add("23_DEVGROUP");
            dtOutputData.Columns.Add("23_AUTHID");
            dtOutputData.Columns.Add("23_DEVGROUPID");
            dtOutputData.Columns.Add("23_DIRECTION");
            dtOutputData.Columns.Add("41_AUTHID");
            dtOutputData.Columns.Add("41_DEVGROUPID");
            dtOutputData.Columns.Add("41_MACID");
            dtOutputData.Columns.Add("41_SHORTNAME");
            dgvOutputFile.DataSource = null;

            for (int j = 0; j < dtOutputFile.Rows.Count; j++)
            {
                dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
                global.Data.logger("CONVERT (phase 3). Processing row #" + j.ToString() + "-----------------", "10001000");    // show # of row in console

                for (int i = 0; i < dtAuth41Data.Rows.Count; i++)
                {
                    string Line = dtAuth41Data.Rows[i].ItemArray[2].ToString().Trim();
                    string[] strArrIn = Line.Split('@');

                    if (dtOutputFile.Rows[j].ItemArray[0].ToString().Trim() == strArrIn[0].Trim()
                        && dtOutputFile.Rows[j].ItemArray[7].ToString().Trim() == dtAuth41Data.Rows[i].ItemArray[15].ToString().Trim()  // check MAC coincides
                        )
                    {
                        DataRow dr = null;
                        dr = dtOutputData.NewRow();
                        dr[0] = dtOutputFile.Rows[j].ItemArray[0].ToString().Trim();
                        dr[1] = dtOutputFile.Rows[j].ItemArray[1].ToString().Trim();
                        dr[2] = dtOutputFile.Rows[j].ItemArray[2].ToString().Trim();
                        dr[3] = dtOutputFile.Rows[j].ItemArray[3].ToString().Trim();
                        dr[4] = dtOutputFile.Rows[j].ItemArray[4].ToString().Trim();
                        dr[5] = dtAuth41Data.Rows[i].ItemArray[0].ToString().Trim();
                        dr[6] = dtOutputFile.Rows[j].ItemArray[6].ToString().Trim();
                        dr[7] = dtOutputFile.Rows[j].ItemArray[7].ToString().Trim();
                        dr[8] = dtAuth41Data.Rows[i].ItemArray[2].ToString().Trim();

                        dtOutputData.Rows.Add(dr);
                    }
                }
            }
            dp.dtSaveToFile(dtOutputData, OUTPUTFILE, true);
            dtOutputFile.Dispose();
            dtOutputData.Dispose();
            dtAuth41Data.Dispose();
        }
        //private void FIND_BIS41_PERSID(string PROCESSOR, string OUTPUTFILE, string DBCONNFILE)      // adds 41_PERSID and 41_VISID as last columns to output file and fills them. Uses LASTNAME, FIRSTNAME and DATEOFBIRTH as input unique data to identify person in 41 
        //{
        //    DataTable dtOutputFile = new DataTable("dtOutputFile");
        //    dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
        //    dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);
        //    dtOutputFile.Columns.Add("41_PERSID");
        //    dtOutputFile.Columns.Add("41_VISID");
        //    dgvOutputFile.DataSource = null;

        //    dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
        //    dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
        //    string connFile = TB_DBCONNFILE_PROCESSOR.Text;
        //    string sqlRqFile = "";
        //    string Data1, Data2, Data3, Data4;
        //    Data1 = Data2 = Data3 = Data4 = "";

        //    for (int j = 0; j < dtOutputFile.Rows.Count; j++)
        //    {
        //        dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
        //        Data1 = dtOutputFile.Rows[j].ItemArray[1].ToString().Trim();              //LASTNAME
        //        Data2 = dtOutputFile.Rows[j].ItemArray[2].ToString().Trim();              //FIRSTNAME
        //        Data3 = dtOutputFile.Rows[j].ItemArray[3].ToString().Trim();              //DATEOFBIRTH 

        //        DataRow rowEdit = dtOutputFile.Rows[j];
        //        rowEdit.BeginEdit();
        //        sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.FIND_BIS41_PERSID_AUTHID_PERSID.sql";
        //        rowEdit["41_PERSID"] = dp.dtFindDB(connFile, sqlRqFile, Data1, Data2, Data3, Data4);            //41_PERSID
        //        sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.FIND_BIS41_PERSID_AUTHID_VISID.sql";
        //        rowEdit["41_VISID"] = dp.dtFindDB(connFile, sqlRqFile, Data1, Data2, Data3, Data4);            //41_VISID
        //        rowEdit.EndEdit();
        //        Data1 = Data2 = Data3 = Data4 = "";
        //    }
        //    dp.dtSaveToFile(dtOutputFile, OUTPUTFILE);
        //}
        //private void FIND_BIS41_PERSID_v2(string PROCESSOR, string OUTPUTFILE, string DBCONNFILE)   // adds 41_PERSID and 41_VISID as last columns to output file and fills them. Uses RESERVE10 (23_PERSID) and RESERVE9 (23_VISID) as input unique data to identify person in 41 
        //{
        //    if (true)
        //        try
        //        {
        //            dataType = "FP";
        //            DataTable dtOutputFile = new DataTable("dtOutputFile");
        //            dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
        //            dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);
        //            dtOutputFile.Columns.Add("41_PERSID");
        //            dtOutputFile.Columns.Add("41_VISID");
        //            dgvOutputFile.DataSource = null;

        //            dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
        //            dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
        //            string connFile = TB_DBCONNFILE_PROCESSOR.Text;
        //            string sqlRqFile = "";
        //            string Data1, Data2, Data3, Data4;
        //            Data1 = Data2 = Data3 = Data4 = "";

        //            for (int j = 0; j < dtOutputFile.Rows.Count; j++)
        //            {
        //                dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
        //                Data1 = dtOutputFile.Rows[j].ItemArray[0].ToString().Trim();                                    //23_PERSID

        //                DataRow rowEdit = dtOutputFile.Rows[j];
        //                rowEdit.BeginEdit();
        //                sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.FIND_BIS41_PERSID.sql";              //41_PERSID (assumed that RESERVE10 has 23_PERSID)
        //                rowEdit["41_PERSID"] = dp.dtFindDB(connFile, sqlRqFile, Data1, Data2, Data3, Data4);
        //                sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.FIND_BIS41_VISID.sql";               //41_VISID  (assumed that RESERVE10 has 23_PERSID)
        //                rowEdit["41_VISID"] = dp.dtFindDB(connFile, sqlRqFile, Data1, Data2, Data3, Data4);
        //                rowEdit.EndEdit();
        //                Data1 = Data2 = Data3 = Data4 = "";
        //                dp.dtSaveToFile(dtOutputFile, OUTPUTFILE);
        //                dtOutputFile.Dispose();
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            msgError = "FIND_BIS41_PERSID_v2. \r\n" +
        //                        dataType + "\r\n" +
        //                        e.Message + "\r\n" +
        //                        "1. Check if all involved files have headers\r\n" +
        //                        "2. Check correspondence of headers and fields in input and associations files" +
        //            "\r\nOptions: Abort the application or Retry to start new operation or Ignore this message";
        //            msgErrorType = "GENERAL EXCEPTION";
        //            global.Data.logger(msgError, "10000100");
        //            DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
        //            if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
        //            if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); }
        //            if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); }
        //        }
        //}
        private void FIND_BIS41_AUTHID(string PROCESSOR, string OUTPUTFILE, string DBCONNFILE)      // adds 41_AUTHID as last column to output file and fills it
        {
            DataTable dtOutputFile = new DataTable("dtOutputFile");
            dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
            DataTable dtOutputData = new DataTable("dtOutputData");
            dtOutputData.Clear(); dtOutputData.Columns.Clear();
            DataTable dtReceivedData = new DataTable("dtOutputData");
            dtReceivedData.Clear(); dtReceivedData.Columns.Clear();

            dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
            dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
            string connFile = TB_DBCONNFILE_PROCESSOR.Text;
            string sqlRqFile = "";
            string Data1 = "";
            string Data2 = "";
            string Data3 = "";
            string Data4 = "";

            dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);

            dtOutputData.Columns.Add("23_PERSID");
            dtOutputData.Columns.Add("LASTNAME");
            dtOutputData.Columns.Add("FIRSTNAME");
            dtOutputData.Columns.Add("DATEOFBIRTH");
            dtOutputData.Columns.Add("23_SHORTNAME");
            dtOutputData.Columns.Add("41_PERSID");
            dtOutputData.Columns.Add("41_VISID");
            //dtOutputData.Columns.Add("41_SHORTNAME");
            //dtOutputData.Columns.Add("41_AUTHID");
            //dtOutputData.Columns.Add("VALIDFROM");
            //dtOutputData.Columns.Add("VALIDUNTIL");
            dtOutputData.Columns.Add("VALIDFROM");
            dtOutputData.Columns.Add("VALIDUNTIL"); 
            dtOutputData.Columns.Add("41_SHORTNAME");
            dtOutputData.Columns.Add("41_AUTHID");
            dgvOutputFile.DataSource = null;

            for (int j = 0; j < dtOutputFile.Rows.Count; j++)
            {
                dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
                DataRow dr = null;
                dr = dtOutputData.NewRow();
                dr[0] = dtOutputFile.Rows[j].ItemArray[0].ToString().Trim();                        //23_PERSID
                dr[1] = Data1 = dtOutputFile.Rows[j].ItemArray[1].ToString().Trim();                //LASTNAME
                dr[2] = Data2 = dtOutputFile.Rows[j].ItemArray[2].ToString().Trim();                //FIRSTNAME
                dr[3] = Data3 = dtOutputFile.Rows[j].ItemArray[3].ToString().Trim();                //DATEOFBIRTH
                dr[4] = Data4 = dtOutputFile.Rows[j].ItemArray[4].ToString().Trim();                //23_SHORTNAME
                dr[5] = dtOutputFile.Rows[j].ItemArray[5].ToString().Trim();                        //41_PERSID
                dr[6] = dtOutputFile.Rows[j].ItemArray[6].ToString().Trim();                        //41_VISID
                //dr[7] = dtOutputFile.Rows[j].ItemArray[7].ToString().Trim();                      //41_SHORTNAME
                //dr[8] = dtOutputFile.Rows[j].ItemArray[8].ToString().Trim();                      //41_AUTHID
                //dr[9] = dtOutputFile.Rows[j].ItemArray[9].ToString().Trim();                      //VALIDFROM
                //dr[10] = dtOutputFile.Rows[j].ItemArray[10].ToString().Trim();                    //VALIDUNTIL
                dr[7] = dtOutputFile.Rows[j].ItemArray[7].ToString().Trim();                        //VALIDFROM
                dr[8] = dtOutputFile.Rows[j].ItemArray[8].ToString().Trim();                        //VALIDUNTIL
                
                dr[9] = dtOutputFile.Rows[j].ItemArray[9].ToString().Trim();                        //41_SHORTNAME
                dr[10] = dtOutputFile.Rows[j].ItemArray[10].ToString().Trim();                      //41_AUTHID
                sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.FIND_BIS41_AUTHID.sql";
                dtReceivedData = dp.dtReadDB(connFile, sqlRqFile, Data1, Data2, Data3, Data4 + "   @%");   //41_AUTHID

                if (dtReceivedData.Rows.Count == 0)
                {
                    global.Data.logger("FIND_BIS41_AUTHID; POSSIBLE TROUBLE; Authorization 23_AUTHID: " + Data4 + " has no 41_AUTHID", "10000100");
                    dtOutputData.Rows.Add(dr);
                }
                else
                {
                    for (int i = 0; i < dtReceivedData.Rows.Count; i++)
                    {
                        dr = null;
                        dr = dtOutputData.NewRow();
                        dr[0] = dtOutputFile.Rows[j].ItemArray[0].ToString().Trim();            //23_PERSID
                        dr[1] = dtOutputFile.Rows[j].ItemArray[1].ToString().Trim();            //LASTNAME
                        dr[2] = dtOutputFile.Rows[j].ItemArray[2].ToString().Trim();            //FIRSTNAME
                        dr[3] = dtOutputFile.Rows[j].ItemArray[3].ToString().Trim();            //DATEOFBIRTH
                        dr[4] = dtOutputFile.Rows[j].ItemArray[4].ToString().Trim();            //23_SHORTNAME
                        dr[5] = dtOutputFile.Rows[j].ItemArray[5].ToString().Trim();            //41_PERSID
                        dr[6] = dtOutputFile.Rows[j].ItemArray[6].ToString().Trim();            //41_VISID
                        dr[7] = dtOutputFile.Rows[j].ItemArray[7].ToString().Trim();            //VALIDFROM
                        dr[8] = dtOutputFile.Rows[j].ItemArray[8].ToString().Trim();            //VALIDUNTIL
                        dr[9] = dtReceivedData.Rows[i].ItemArray[0].ToString().Trim();          //41_SHORTNAME
                        dr[10] = dtReceivedData.Rows[i].ItemArray[1].ToString().Trim();         //41_PERSID
                        dtOutputData.Rows.Add(dr);
                    }
                    global.Data.logger("CONVERT (phase 3). FIND_BIS41_AUTHID: processing row #" + j.ToString() + "", "10001000");    // show # of row in console
                    global.Data.logger("Authorization: " + Data4 + " is assigned to: \r\n" + Data1 + " " + Data2 + " " + Data3, "10001000");
                }
            }
            dp.dtSaveToFile(dtOutputData, OUTPUTFILE, true);
            dtOutputFile.Dispose();
            dtOutputData.Dispose();
            dtReceivedData.Dispose();
        }
        private void FIND_BIS41_COMPANYID_EMPLOYEES(string PROCESSOR, string OUTPUTFILE, string AUX_FILE, string DBCONNFILE) // adds 41_COMPANYID as last column to output file and fills it
        {
            if (true)
                try
                {
                    DataTable dtOutputFile = new DataTable("dtOutputFile");
                    dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
                    dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);
                    dtOutputFile.Columns.Add("41_COMPANYID");                                               //41_COMPANYID
                    //dtOutputFile.Columns.Add("41_COMPANYNO");                                               //41_COMPANYNO

                    DataTable dtCompany = new DataTable("dtCompany");
                    dtCompany.Clear(); dtCompany.Columns.Clear();
                    dtCompany = dp.dtReadFromFile(AUX_FILE, true);                                                                          // read file with Companies data. 

                    dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
                    dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
                    string Data1, Data2, Data3, Data4;
                    Data1 = Data2 = Data3 = Data4 = "";
                    string connFile = TB_DBCONNFILE_PROCESSOR.Text;
                    string sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.FIND_BIS41_COMPANYID.sql";

                    for (int j = 0; j < dtOutputFile.Rows.Count; j++)
                    {
                        dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
                        global.Data.logger("CONVERT (phase 3). FIND_BIS41_COMPANYID_EMPLOYEES: processing row #" + j.ToString() + "", "10001000");    // show # of row in console

                        for (int i = 0; i < dtCompany.Rows.Count; i++)
                        {
                            if (dtOutputFile.Rows[j].ItemArray[46].ToString().Trim() == dtCompany.Rows[i].ItemArray[0].ToString().Trim())   //get 23_COMPANYID from Output file data (field46)
                            {
                                Data1 = dtCompany.Rows[i].ItemArray[1].ToString().Trim();                                                   //get 23_COMPANYNO (unique field) from Companies file data (on known COMPANYID)
                                break;
                            }
                        }

                        DataRow rowEdit = dtOutputFile.Rows[j];
                        string rowData = dp.dtFindDB(connFile, sqlRqFile, Data1, Data2, Data3, Data4);                                      //get 41_COMPANYID (on known COMPANYNO) from target system and put it into the last column
                        rowEdit.BeginEdit();
                        rowEdit["41_COMPANYID"] = rowData;
                        rowEdit.EndEdit();
                        global.Data.logger(
                            "CONVERT (phase 3). dtFindDB----- \r\n" +
                            "connFile: ..." + connFile.Substring(connFile.Length - visibleFileName, visibleFileName) + "\r\n" +
                            "sqlRqFile: ..." + sqlRqFile.Substring(sqlRqFile.Length - visibleFileName, visibleFileName) + "\r\n" +
                            "Data1: " + Data1 + "\r\n" +
                            "Data2: " + Data2 + "\r\n" +
                            "found: " + rowData + "\r\n"
                            , "10001000");
                        Data1 = "";
                    }
                    dp.dtSaveToFile(dtOutputFile, OUTPUTFILE, true);
                    dtOutputFile.Dispose();
                    dtCompany.Dispose();
                    //ReadOutputFile(OUTPUTFILE);
                }
                catch { }
        }
        private void FIND_BIS41_COMPANYID_VISITORS(string PROCESSOR, string OUTPUTFILE, string AUX_FILE, string DBCONNFILE)
        {
        if (true)
            try
            {
                DataTable dtOutputFile = new DataTable("dtOutputFile");
                dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
                dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);
                dtOutputFile.Columns.Add("41_COMPANYID");                                               //41_COMPANYID
                //dtOutputFile.Columns.Add("41_COMPANYNO");                                               //41_COMPANYNO

                DataTable dtCompany = new DataTable("dtCompany");
                dtCompany.Clear(); dtCompany.Columns.Clear();
                dtCompany = dp.dtReadFromFile(AUX_FILE, true);

                dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
                dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);
                string Data1, Data2, Data3, Data4;
                Data1 = Data2 = Data3 = Data4 = "";
                string connFile = TB_DBCONNFILE_PROCESSOR.Text;
                string sqlRqFile = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "FP.FIND_BIS41_COMPANYID.sql";

                for (int j = 0; j < dtOutputFile.Rows.Count; j++)
                {
                    dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
                    global.Data.logger("CONVERT (phase 3). FIND_BIS41_COMPANYID_VISITORS: processing row #" + j.ToString() + "", "10001000");    // show # of row in console

                    for (int i = 0; i < dtCompany.Rows.Count; i++)
                    {
                        if (dtOutputFile.Rows[j].ItemArray[57].ToString().Trim() == dtCompany.Rows[i].ItemArray[0].ToString().Trim())
                        {
                            Data1 = dtCompany.Rows[i].ItemArray[1].ToString().Trim();
                            break;
                        }
                    }

                    DataRow rowEdit = dtOutputFile.Rows[j];
                    string rowData = dp.dtFindDB(connFile, sqlRqFile, Data1, Data2, Data3, Data4);  
                    rowEdit.BeginEdit();
                    rowEdit["41_COMPANYID"] = rowData;                                                                          //41_COMPANYID
                    rowEdit.EndEdit();
                    global.Data.logger(
                        "CONVERT (phase 3). dtFindDB----- \r\n" +
                        "connFile: ..." + connFile.Substring(connFile.Length - visibleFileName, visibleFileName) + "\r\n" +
                        "sqlRqFile: ..." + sqlRqFile.Substring(sqlRqFile.Length - visibleFileName, visibleFileName) + "\r\n" +
                        "Data1: " + Data1 + "\r\n" +
                        "Data2: " + Data2 + "\r\n" +
                        "found: " + rowData + "\r\n"
                        , "10001000");
                    Data1 = "";
                }
                dp.dtSaveToFile(dtOutputFile, OUTPUTFILE, true);
                dtOutputFile.Dispose();
                //ReadOutputFile(OUTPUTFILE);
            }
            catch { }
        }
        private void FILL_CARD_ATTRIBUTES(string PROCESSOR, string OUTPUTFILE, string DBCONNFILE)
        {
            //FIND_BIS41_PERSID_v2(PROCESSOR, OUTPUTFILE, DBCONNFILE);
            DataTable dtOutputFile = new DataTable("dtOutputFile");
            dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
            dtOutputFile = dp.dtReadFromFile(OUTPUTFILE, true);
            dtOutputFile.Columns.Add("41_CARDID");                                      //41_CARDID, Data0
            dtOutputFile.Columns.Add("41_CODEDATA");                                    //41_CODEDATA, Data3
            dtOutputFile.Columns.Add("41_DATECREATED");                                 //41_DATECREATED, Data4
            dtOutputFile.Columns.Add("41_CLIENTID");                                    //41_CLIENTID, Data5
            dtOutputFile.Columns.Add("41_USAGETYPEID");                                 //41_DATECREATED, Data6
            dgvOutputFile.DataSource = null;
            string msg = "FP: FILL_CARD_ATTRIBUTES";

            dp_updateProgressBarsMaximum(this, "CONVERT", -1, dtOutputFile.Rows.Count, -1);
            dp_updateProgressBarsCurrent(this, "CONVERT", -1, 0, -1);

            //if (lblLoginStatus.Text != "ONLINE") APILogin(TB_API_LOGIN_SERVER.Text, API_LOGIN_NAME.Text, API_LOGIN_PWD.Text);
            if (lblLoginStatus.Text == "ONLINE")
            {
                // Add test card to get proper data of fields of DB
                global.Data.logger(msg + ": PREPARING TEST CARD TO BE ADDED", "10000010");
                int CARDNO = 999999; int CODEDATA = 234; string aceCardID = "";                                                                                                //Prepare data of test card
                aceCardID = dp.dtFindDB(TB_DBCONNFILE_OUT.Text, TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "bis41.extRQ.06.sql", CARDNO.ToString("D12"), CODEDATA.ToString(), "", ""); //Check if 41_CARDID of of test card exists

                if (aceCardID == "")
                {
                    ACECards aceCard = new ACECards(ace);
                    aceCard.CARDNO = CARDNO.ToString("D12");
                    aceCard.CODEDATA = CODEDATA.ToString();
                    result = aceCard.Add();
                    if (result != API_RETURN_CODES_CS.API_SUCCESS_CS) MessageBox.Show("Can't add test card. Restart target system!");
                    else aceCardID = dp.dtFindDB(TB_DBCONNFILE_OUT.Text, TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "bis41.extRQ.06.sql", CARDNO.ToString("D12"), CODEDATA.ToString(), "", "");      //Get 41_CARDID of added test card
                    if (aceCard.Get(aceCardID) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                    {
                        global.Data.logger(msg + ": TEST CARD ADDED OK", "10000010");     
                        dtTmp.Rows.Clear(); dtTmp.Clear(); dtTmp.Columns.Clear();
                            dtTmp = dp.dtReadDB(TB_DBCONNFILE_OUT.Text, TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "bis41.extRQ.10.sql", aceCardID, "", "", "");      //Get 41_CARDID of added test card

                            for (int j = 0; j < dtOutputFile.Rows.Count; j++)                                                       //Add all the cards
                            {
                                dp_updateProgressBarsCurrent(this, "CONVERT", -1, j + 1, -1);
                                                                             
                                string hexL = Convert.ToInt32(dtOutputFile.Rows[j].ItemArray[5].ToString().Trim()).ToString("X8"); //CARDNO
                                string hexH = Convert.ToInt32(dtOutputFile.Rows[j].ItemArray[6].ToString().Trim()).ToString("X8"); //CODEDATA

                                DataRow rowEdit = dtOutputFile.Rows[j];
                                rowEdit.BeginEdit();
                                rowEdit["41_CARDID"] = "0012FF" + j.ToString("X10");
                                rowEdit["41_CODEDATA"] = "0x" + hexH + hexL;
                                rowEdit["41_DATECREATED"] = "2016-01-01 00:00:00.000";                                              //DATECREATED   "2016-01-01 00:00:00.000"; dtTmp.Rows[0].ItemArray[8].ToString();
                                rowEdit["41_CLIENTID"] = dtTmp.Rows[0].ItemArray[18].ToString();                                    //CLIENTID      "FF0000D300000002"; <- the same as test card has
                                rowEdit["41_USAGETYPEID"] = dtTmp.Rows[0].ItemArray[21].ToString();                                 //USAGETYPEID   "FF00011A00000001"; <- the same as test card has
                                rowEdit.EndEdit();
                                global.Data.logger(msg + ": " + j.ToString(), "10001000");                                          // show processed # of line in console
                            }
                            result = aceCard.Delete(); if (result == API_RETURN_CODES_CS.API_SUCCESS_CS) global.Data.logger(msg + ": TEST CARD DELETED OK", "10000010");
                            dp.dtSaveToFile(dtOutputFile, OUTPUTFILE, true);
                            dtOutputFile.Dispose();
                    }
                }
                else
                {
                    msgError = "FIND_BIS41_PERSID_FILL_CARD_ATTRIBUTES. \r\n" +
                                "CanNOT add test card. Such undeleted card already exists\r\n" +
                                "1. Delete card with N999999-234 and restart conversion once again";
                    msgErrorType = "API OFFLINE";
                    global.Data.logger(msgError, "10000100");
                    MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                msgError = "FIND_BIS41_PERSID_FILL_CARD_ATTRIBUTES. \r\n" +
                            "ACE API must be online to finish operation correctly\r\n" +
                            "1. Check ACE API connection and restart conversion once again";
                msgErrorType = "API OFFLINE";
                global.Data.logger(msgError, "10000100");
                MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        
#endregion ------------------------------------------------------------------------------

#region SEND DATA TO TARGET SYSTEM ---------------------------------------------------------------------------------

        private void cbb_Job_SelectedIndexChanged(object sender, EventArgs e)
        {
            //controlsStatus_Blocked();
            jobOption = cbb_Job.SelectedItem.ToString();
            ReadDataFromJoblist();
        }
        private void ReadDataFromJoblist()
        {
            if (true) 
                try
            {   
            dataType = "NEW JOB SELECTED";
            dp_updateProgressBarsCurrent(this, "cbb_Job_SelectedIndexChanged", 0, 0, 0);
            dp_updateProgressBarsMaximum(this, "cbb_Job_SelectedIndexChanged", 100, 100, 100);

            DataTable dtJoblist = new DataTable("dtJoblist");
            dtJoblist.Rows.Clear(); dtJoblist.Clear(); dtJoblist.Columns.Clear();
            string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JOBLIST.Text;
            dtJoblist = loadJobList(jobListPath, false);
            dgv_JOBLIST.DataSource = dtJoblist;

            for (int j = 0; j < dtJoblist.Rows.Count; j++)
            {
                if (jobOption == dtJoblist.Rows[j].ItemArray[0].ToString().Trim())
                {
                    this.TB_OUTPUT_TABLE.Text = dtJoblist.Rows[j].ItemArray[1].ToString().Trim();
                    this.TB_INPUTFILE.Text = TB_WORKFILESPATH.Text + @"\_dataIn\" + dtJoblist.Rows[j].ItemArray[2].ToString().Trim();
                    this.TB_DBCONNFILE_IN.Text = TB_WORKFILESPATH.Text + @"\_dbConnections\" + dtJoblist.Rows[j].ItemArray[3].ToString().Trim();
                    this.TB_SQLFILE_IN.Text = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + dtJoblist.Rows[j].ItemArray[4].ToString().Trim();
                    this.TB_APIASSOCFILE.Text = TB_WORKFILESPATH.Text + @"\_associations\" + dtJoblist.Rows[j].ItemArray[5].ToString().Trim();
                    this.TB_OUTPUTFILE.Text = TB_WORKFILESPATH.Text + @"\_dataOut\" + dtJoblist.Rows[j].ItemArray[6].ToString().Trim();
                    this.TB_APISCHEMAFILE.Text = TB_WORKFILESPATH.Text + @"\_schemas\" + dtJoblist.Rows[j].ItemArray[7].ToString().Trim();
                    this.lblJobDescription.Text = dtJoblist.Rows[j].ItemArray[8].ToString().Trim();
                    this.TB_DBCONNFILE_PROCESSOR.Text = TB_WORKFILESPATH.Text + @"\_dbConnections\" + dtJoblist.Rows[j].ItemArray[9].ToString().Trim();
                    this.TB_PROCESSOR.Text = dtJoblist.Rows[j].ItemArray[10].ToString().Trim();
                    this.TB_DBCONNFILE_OUT.Text = TB_WORKFILESPATH.Text + @"\_dbConnections\" + dtJoblist.Rows[j].ItemArray[11].ToString().Trim();
                    this.TB_SQLFILE_OUT.Text = TB_WORKFILESPATH.Text + @"\_sqlQueries\" + dtJoblist.Rows[j].ItemArray[12].ToString().Trim();
                    this.cmb_Output.Text = dtJoblist.Rows[j].ItemArray[13].ToString().Trim();
                    this.TB_AUX_FILE.Text = TB_WORKFILESPATH.Text + @"\_dataIn\" + dtJoblist.Rows[j].ItemArray[14].ToString().Trim();
                    if (String.IsNullOrEmpty(dtJoblist.Rows[j].ItemArray[15].ToString().Trim())) cb_FILE_PROCESSOR.Checked = false; else cb_FILE_PROCESSOR.Checked = true;
                }
            }

            dgvInputFile.DataSource =
            dgvOutputFile.DataSource =
            dgvAssocFile.DataSource =
            dgvAPISchemaFile.DataSource =
            dgvDisplayOutput.DataSource =
            null;
            //dgvInputFile.Rows.Clear(); dgvInputFile.Columns.Clear();
            //dgvOutputFile.Rows.Clear(); dgvOutputFile.Columns.Clear();
            //dgvAssocFile.Rows.Clear(); dgvAssocFile.Columns.Clear();
            //dgvAPISchemaFile.Rows.Clear(); dgvAPISchemaFile.Columns.Clear();
            //dgvDisplayOutput.Rows.Clear(); dgvDisplayOutput.Columns.Clear();

            //ReadInputFile(this.TB_INPUTFILE.Text, cb_Headers.Checked);
            dgvAssocFile.DataSource = ReadAssocFile(this.TB_APIASSOCFILE.Text);
            //ReadOutputFile(this.TB_OUTPUTFILE.Text);
            dgvAPISchemaFile.DataSource = ReadApiSchemaFile(this.TB_APISCHEMAFILE.Text);

            TB_INPUTFILE_ROWS.Text =
            TB_INPUTFILE_COLUMNS.Text =
            TB_OUTPUTFILE_ROWS.Text =
            TB_OUTPUTFILE_COLUMNS.Text =
                //TB_ASSOCFILE_ROWS.Text          =
                //TB_ASSOCFILE_COLUMNS.Text       =
                //TB_APISCHEMAFILE_ROWS.Text      =
                //TB_APISCHEMAFILE_COLUMNS.Text   =
            "";
            TB_API_PROCESSOR_STARTLINE.Text =
            TB_API_PROCESSOR_LINES.Text =
            "0";
            if (jobOption == "t01. Пробная загрузка  Сотрудников") btn_FullAuto.Visible = true; else btn_FullAuto.Visible = false;

        }
                catch (Exception ex)
                {
                    msgError = "Job_SelectedIndexChanged. \r\n" +
                                dataType + "\r\n" +
                                ex.Message + "\r\n" +
                                "JOB: BAD OPTIONS\r\n" +
                                "1. Check joblist for missed parameters (missed column data in lines?)\r\n" +
                                "2. Check if 'joblist' file has proper structure\r\n" +
                    "Options: Abort the application or Retry to start new operation or Ignore this message";
                    msgErrorType = "GENERAL EXCEPTION";
                    global.Data.logger(msgError, "10000100");
                    DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                    if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); }
                    if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); }
                }
        }
        private void btn_Start_Click(object sender, EventArgs e)
        {
            sendMethod = cmb_Output.SelectedItem.ToString();
            switch (sendMethod)
            {
                case "ACE API Cmd":
                    if (cb_TestRun.Checked || F_API_ONLINE) SendData(TB_OUTPUT_TABLE.Text, sendMethod);     // Check if API is online
                    if (!cb_TestRun.Checked && !F_API_ONLINE)                                               // If API is not online try to Login several times
                    {
                        for (int j = 1; j <= Convert.ToInt32(TB_ATTEMPTS.Text); j++)
                        {
                            if (APILogin(TB_API_LOGIN_SERVER.Text, API_LOGIN_NAME.Text, API_LOGIN_PWD.Text))
                            {
                                F_API_ONLINE = true;
                                global.Data.logger("API WAS OFFLINE; AUTOLOGIN ENABLED; ", "10000011");
                                break;
                            }
                            else
                            {
                                F_API_ONLINE = false;
                                global.Data.logger("API IS OFFLINE; AUTOLOGIN ATTEMPT; " + j.ToString(), "10000011");
                            }
                        }

                        if (F_API_ONLINE) SendData(TB_OUTPUT_TABLE.Text, sendMethod);                           // If after autologin API is online then send data the normal way
                        if (!F_API_ONLINE && rb_ShowMsg.Checked) MessageBox.Show("API IS ULTIMATELY OFFLINE!"); // if after autologin API is still offline then show message
                        if (!F_API_ONLINE && rb_Scenario1.Checked)                                              // ... or try to reboot API server   
                        {
                            Process myProcess = new Process();
                            myProcess.StartInfo.FileName = TB_WORKFILESPATH.Text + @"\_cmd\" + TB_CMD_FILE.Text;
                            //myProcess.StartInfo.Arguments = @"/C cd " + Application.StartupPath + "/server/login & start.bat";
                            //myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            //myProcess.StartInfo.CreateNoWindow = true;
                            myProcess.Start();
                            global.Data.logger("API IS DEAD; EXTERNAL CMD-FILE EXECUTED; " + TB_CMD_FILE.Text, "10001111");

                                
                            for (int i = 0; i <= Convert.ToInt32(TB_WAIT_RESTART.Text); i++)
                            {
                                global.Data.logger("SLEEPING...; TIMER; " + i.ToString() + " from " + Convert.ToInt32(TB_WAIT_RESTART.Text) + " seconds", "10000000");
                                System.Threading.Thread.Sleep(1000);
                            }
                            global.Data.logger("API IS DEAD; RESUMING JOB AFTER RESTART; ", "10001111");

                            if (APILogin(TB_API_LOGIN_SERVER.Text, API_LOGIN_NAME.Text, API_LOGIN_PWD.Text))
                            {
                                F_API_ONLINE = true;
                                SendData(TB_OUTPUT_TABLE.Text, sendMethod);
                            }
                            else
                            {
                                msgError = "API IS ULTIMATELY DEAD; RESUMING JOB IS IMPOSSIBLE; ";
                                global.Data.logger(msgError, "10001111");
                                MessageBox.Show(msgError);
                            }
                        }
                    }
                    break;
                case "T-SQL Script": 
                    SendData(TB_OUTPUT_TABLE.Text, sendMethod);
                    break;
                case "CS":
                    SendData(TB_OUTPUT_TABLE.Text, sendMethod); 
                    break;
                default: break;
            }
        }
        private void SendData(string dataType, string sendMethod)
        {
            if (cb_AutoRun.Checked)
            {
                dtOutputFile = ReadOutputFile(this.TB_OUTPUTFILE.Text);                                     // read data from source files
                dtAPISchemaFile = ReadApiSchemaFile(this.TB_APISCHEMAFILE.Text);

                int REC_N_Start = Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text);
                int REC_N_End = Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text) + Convert.ToInt32(this.TB_API_PROCESSOR_LINES.Text) - 1;

                if (cb_StatAutoReset.Checked) ResetStatistics();

                for (int REC_N = REC_N_Start; REC_N <= REC_N_End; REC_N++)
                {
                    if (sendMethod == "ACE API Cmd")    if (!API_SendData_Single(dtOutputFile, dtAPISchemaFile, dataType, REC_N)) break;
                    if (sendMethod == "T-SQL Script")   if (!SQL_SendData_Single(dtOutputFile, dtAPISchemaFile, dataType, REC_N)) break;
                    if (sendMethod == "CS")             if (!CS_SendData_Single(dtOutputFile, dtAPISchemaFile, dataType, REC_N)) break;
                    
                    dp_updateProgressBarsCurrent(this, "SEND. Autorun ", -1, -1, pb_SendData.Value + 1);    // update progress bar
                }
                
                if (sendMethod == "ACE API Cmd" && cb_LogRAM.Checked) dp.dumpRAMLog(1);                     // dump RAM Log to HDD if RAM logging was enabled
                cb_AutoRun.Checked = false;
                TB_API_PROCESSOR_STARTLINE.Text = "0";
                TB_API_PROCESSOR_LINES.Text = "0";
                dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
            }

            if (cb_Single.Checked)
            {
                if (cb_LogRAM.Checked) cb_LogRAM.Checked = false;                                       // for single step operation mode RAM logging makes no sence
                
                if (global.Data.REC_N == Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text))         // only for the first record: read data from source files
                {
                    dtOutputFile = ReadOutputFile(this.TB_OUTPUTFILE.Text);                             // read data from source files
                    dtAPISchemaFile = ReadApiSchemaFile(this.TB_APISCHEMAFILE.Text);
                    if (cb_StatAutoReset.Checked) ResetStatistics();
                }

                lbl_Current.Text = global.Data.REC_N.ToString();
                if (sendMethod == "ACE API Cmd")    API_SendData_Single(dtOutputFile, dtAPISchemaFile, dataType, global.Data.REC_N);
                if (sendMethod == "T-SQL Script")   SQL_SendData_Single(dtOutputFile, dtAPISchemaFile, dataType, global.Data.REC_N);
                if (sendMethod == "CS")             CS_SendData_Single(dtOutputFile,  dtAPISchemaFile, dataType, global.Data.REC_N);
                global.Data.REC_N++;
                dp_updateProgressBarsCurrent(this, "SEND. Single ", -1, -1, pb_SendData.Value + 1);
                //dgvOutputFile.Rows[global.Data.REC_N].Cells[0].Selected = true;                       // highlight current cell in DataGridView
                //dgvOutputFile.CurrentCell = dgvOutputFile.Rows[global.Data.REC_N].Cells[0];           // highlight current cell in DataGridView
                if (global.Data.REC_N == Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text) + Convert.ToInt32(this.TB_API_PROCESSOR_LINES.Text))     // if last record is processed then stop
                {
                    cb_Single.Checked = false;
                    TB_API_PROCESSOR_STARTLINE.Text = "0";
                    TB_API_PROCESSOR_LINES.Text = "0";
                    dtOutputFile.Clear(); dtOutputFile.Columns.Clear();
                }
            }
        }
        private bool displayOutputData(DataTable dtOutputFile, DataTable dtAPISchemaFile, int REC_N)
        {
            if (true)
                try
                {
                    dataType = "Preparing output data";
                    dtOutputData.Rows.Clear(); dtOutputData.Clear(); dtOutputData.Columns.Clear();
                    dtOutputData.Columns.Add("N");
                    dtOutputData.Columns.Add("PARAMETER");
                    dtOutputData.Columns.Add("DATA");
                    dgvDisplayOutput.DataSource = null;

                    for (int i = 0; i < dtAPISchemaFile.Rows.Count; i++)
                    {
                        if (dtAPISchemaFile.Rows[i].ItemArray[2].ToString() != "")
                        {
                            string Data = dtOutputFile.Rows[REC_N].ItemArray[Convert.ToInt32(dtAPISchemaFile.Rows[i].ItemArray[2].ToString())].ToString();
                            dtOutputData.Rows.Add(dtAPISchemaFile.Rows[i].ItemArray[0].ToString(), dtAPISchemaFile.Rows[i].ItemArray[1].ToString(), Data);
                        }
                        else dtOutputData.Rows.Add(dtAPISchemaFile.Rows[i].ItemArray[0].ToString(), dtAPISchemaFile.Rows[i].ItemArray[1].ToString(), "");
                    }
                    if (!cb_AutoRun.Checked && cb_DisplayData_Send.Checked) dgvDisplayOutput.DataSource = dtOutputData;
                    return true;
                }
                catch (Exception e)
                {
                    msgError = "displayOutputData. \r\n" +
                                dataType + "\r\n" +
                                e.Message + "\r\n" +
                                "1. Check if output file was converted properly\r\n2. Check correspodnence between fields of output file and schema file" + "\r\n" +
                                "Options: Abort the application or Retry to start new operation or Ignore this message";
                    msgErrorType = "GENERAL EXCEPTION";
                    global.Data.logger(msgError, "10000100");
                    DialogResult      = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    if (DialogResult == DialogResult.Abort)     { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                    if (DialogResult == DialogResult.Retry)     { global.Data.logger(DialogResult.ToString(), "10000100"); return false; }
                    if (DialogResult == DialogResult.Ignore)    { global.Data.logger(DialogResult.ToString(), "10000100"); return true; }
                }
            return false;
        }

        private bool API_SendData_Single(DataTable dtOutputFile, DataTable dtAPISchemaFile, string dataType, int REC_N)
        {
            #region API. Define common variables ------------------------------------------------------------------------

//dp.logger("cp03: API_SendData_Single: API. Define common variables", "11000100"); 
            string  aceID           = "";
            string  acePersonID     = "";
            string  aceCardID       = "";
            string  aceAuthID       = "";
            string  aceVisitorID    = "";
            bool    F_RECORD_EXISTS = false;                                            // flag 'Record exists'     (ID is found in database)
            bool    F_VISITOR       = false;                                            // flag 'Visitor'           (not Person)
            bool    F_PERSON        = false;                                            // flag 'Person'
                    F_PROCESSED     = false;
            string  RESULT          = REC_N.ToString() + "; " + dataType;
            string  dateDefault     = "20200101";
            string  dtFormat        = "yyyyMMdd";                                       //string maximal dtFormat = "yyyy-MM-dd HH:mm:ss.000";
            string  value           = "";
            DateTime DTime          = DateTime.ParseExact(dateDefault, dtFormat, null);
            ACEAuthorizations   aceAuthorization    = null;
            ACEPersons          acePerson           = null;
            ACEVisitors         aceVisitor          = null;
            ACECompanies        aceCompany          = null;
            ACECards            aceCard             = null;
            ACETimeModels       aceTimeModel        = null;
            ACEDateT            aceDateOfBirth      = null;
            ACEDateT            aceIdValidUntil     = null;
            ACEDateT            aceDate             = null;
            ACEDateT            aceDateFrom         = null;
            ACEDateT            aceDateTill         = null;
            ACEBoolT            flag                = null;
            //dtFormat = TB_DAYTIME_FORMAT.Text;

#endregion--------------

            if (displayOutputData(dtOutputFile, dtAPISchemaFile, REC_N))                            //Prepare and display output data. If output data is ready then continue (assign it to api variables)
                try
                {
                #region API. COMPANIES ------------------------------------------------------------------------------
                if (dataType == "COMPANIES")
                {
                    
                        //RESULT = dataType;
                        aceCompany = new ACECompanies(ace);
                        for (int i = 0; i < dtOutputData.Rows.Count; i++)
                        {
                            value = dtOutputData.Rows[i].ItemArray[2].ToString();
                            if (!String.IsNullOrEmpty(value) && value != "NULL")
                                switch (i)
                                {
                                    case 0: aceID = value;                                                                          //COMPANYID
                                        if (aceCompany.Get(aceID) == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) F_API_ONLINE = false; 
                                        if (aceCompany.Get(aceID) == API_RETURN_CODES_CS.API_SUCCESS_CS) F_RECORD_EXISTS = true; 
                                        break;                             
                                    case 1: aceCompany.COMPANYNO = value; break;
                                    case 2: aceCompany.NAME = value; break;
                                    case 3: aceCompany.REMARKS = value; break;
                                    case 4: aceCompany.STREETHOUSENO = value; break;
                                    case 5: aceCompany.EMAIL = value; break;
                                    case 6: aceCompany.CITY = value; break;
                                    case 7: aceCompany.PHONE = value; break;
                                    case 8: aceCompany.FAX = value; break;
                                    case 9: aceCompany.MOBILEPHONE = value; break;
                                    case 10: aceCompany.HOMEPAGE = value; break;
                                    case 11: aceCompany.ZIPCODE = value; break;
                                    default: break;
                                }
                        }
                    }

                #endregion-----------------------------------------------------------------------------
                #region API. TIME MODELS ------------------------------------------------------------------------------
                if (dataType == "TIME MODELS")
                    {
                        //RESULT = dataType;
                        aceTimeModel = new ACETimeModels(ace);
                        aceDate = new ACEDateT();
                        flag = new ACEBoolT();
                        for (int i = 0; i < dtOutputData.Rows.Count; i++)
                        {
                            value = dtOutputData.Rows[i].ItemArray[2].ToString();
                            if (!String.IsNullOrEmpty(value) && value != "NULL")
                                switch (i)
                                {
                                    case 0: aceID = value;                                                                                              //TMID
                                        if (aceTimeModel.Get(aceID) == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) F_API_ONLINE = false; 
                                        if (aceTimeModel.Get(aceID) == API_RETURN_CODES_CS.API_SUCCESS_CS) F_RECORD_EXISTS = true; 
                                        break;                             
                                    case 1: aceTimeModel.NAME = value; break;
                                    case 2: aceTimeModel.DESCRIPTION = value; break;
                                    case 3: aceDate.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); aceTimeModel.REFDATE = aceDate; break;
                                    case 4: flag.op_Assign(false); if (value != "0") flag.op_Assign(true); aceTimeModel.IGNORESPECDAYS = flag; ; break;
                                    default: break;
                                }
                        }
                    }

                #endregion--------------------------------------------------------------------------------
                #region API. EMPLOYEES ------------------------------------------------------------------------------
                if (dataType == "EMPLOYEES") 
                    {
                        acePerson = new ACEPersons(ace);
                        for (int i = 0; i < dtOutputData.Rows.Count; i++)
                        {
                            value = dtOutputData.Rows[i].ItemArray[2].ToString();
                            if (!String.IsNullOrEmpty(value) && value != "NULL")
                                switch (i)
                                {                                                                                                   // table [PERSONS]  - begin
                                    case 0: aceID = value;
                                        if (acePerson.Get(aceID) == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) F_API_ONLINE = false; 
                                        if (acePerson.Get(aceID) == API_RETURN_CODES_CS.API_SUCCESS_CS) F_RECORD_EXISTS = true; 
                                        break;
                                    case 1: acePerson.LASTNAME = value; break;
                                    case 2: acePerson.FIRSTNAME = value; break;
                                    case 3: acePerson.ADDITIONALLASTNAME = value; break;
                                    case 4: aceDateOfBirth = new ACEDateT(); aceDateOfBirth.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); acePerson.DATEOFBIRTH = aceDateOfBirth; break; //"yyyyMMdd"
                                    case 5: acePerson.PERSNO = value; break;
                                    case 6: acePerson.SEX = (ACESexT)Convert.ToInt32(value); break;
                                    case 7: acePerson.GRADE = value; break;
                                    case 8: acePerson.NUMBERPLATE = value; break;
                                    case 9: acePerson.STREETHOUSENO = value; break;
                                    case 10: acePerson.ZIPCODE = value; break;
                                    case 11: acePerson.CITY = value; break;
                                    case 12: acePerson.COUNTRY = value; break;
                                    case 13: acePerson.NATIONALITY = value; break;
                                    case 14: acePerson.PHONEPRIVATE = value; break;
                                    case 15: acePerson.FAXPRIVATE = value; break;
                                    case 16: acePerson.PHONEOFFICE = value; break;
                                    case 17: acePerson.FAXOFFICE = value; break;
                                    case 18: acePerson.PHONEMOBILE = value; break;
                                    case 19: acePerson.PHONEOTHER = value; break;
                                    case 20: acePerson.EMAIL = value; break;
                                    case 21: acePerson.WEBPAGEURL = value; break;
                                    case 22: acePerson.MAIDENNAME = value; break;
                                    case 23: acePerson.CITYOFBIRTH = value; break;
                                    case 24: acePerson.MARITALSTATUS = (ACEMaritalStateT)Convert.ToInt32(value); break;
                                    case 25: acePerson.IDTYPE = value; break;
                                    case 26: acePerson.IDNUMBER = value; break;
                                    case 27: aceIdValidUntil = new ACEDateT(); aceIdValidUntil.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); acePerson.IDVALIDUNTIL = aceIdValidUntil; break; //"yyyyMMdd"
                                    case 28: acePerson.HEIGHT = Convert.ToInt32(value); break;
                                    case 29: acePerson.DEPARTMENT = value; break;
                                    case 30: acePerson.CENTRALOFFICE = value; break;
                                    case 31: acePerson.COSTCENTRE = value; break;
                                    case 32: acePerson.JOB = value; break;
                                    case 33: acePerson.ATTENDANT = value; break;
                                    case 34: acePerson.REASONSTAY = value; break;
                                    case 35: acePerson.REMARK = value; break;
                                    case 36: acePerson.RESERVE1 = value; break;
                                    case 37: acePerson.RESERVE2 = value; break;
                                    case 38: acePerson.RESERVE3 = value; break;
                                    case 39: acePerson.RESERVE4 = value; break;
                                    case 40: acePerson.RESERVE5 = value; break;
                                    case 41: acePerson.RESERVE6 = value; break;
                                    case 42: acePerson.RESERVE7 = value; break;
                                    case 43: acePerson.RESERVE8 = value; break;
                                    case 44: acePerson.RESERVE9 = value; break;
                                    case 45: acePerson.RESERVE10 = value; break;                                         // table [PERSONS]  - end
                                    case 46: acePerson.COMPANYID = value; break; //41_COMPANYID
                                    //case 47: acePerson.IDENTPAPER = value; break;
                                    //case 48: acePerson.PASSPORTNO = value; break;
                                    //case 49: acePerson.REASON = value; break;
                                    //case 50: acePerson.LOCATION = value; break;
                                    //case 51: acePerson.REMARKS = value; break;
                                    //case 52: break;                         //CARDID
                                    //case 53: DTime = new DateTime(); DTime = DateTime.ParseExact(value, "yyyyMMdd", null); acePerson.AUTHFROM = DTime; break;
                                    //case 54: DTime = new DateTime(); DTime = DateTime.ParseExact(value, "yyyyMMdd", null); acePerson.AUTHUNTIL = DTime; break;
                                    //case 55: ACEDateT ARRIVALEXPECTED = new ACEDateT(); ARRIVALEXPECTED.Set(dp.DAY(value, dtFormat), dp.MONTH(value, dtFormat), dp.YEAR(value, dtFormat)); acePerson.ARRIVALEXPECTED = ARRIVALEXPECTED; break;
                                    //case 56: ACEDateT DEPARTEXPECTED = new ACEDateT(); DEPARTEXPECTED.Set(dp.DAY(value, dtFormat), dp.MONTH(value, dtFormat), dp.YEAR(value, dtFormat)); acePerson.DEPARTEXPECTED = DEPARTEXPECTED; break;
                                    default: break;                                                                                  // table [VISITORS] - end
                                }
                        }
                    }

                #endregion-------------------------------------------------------------------------------------
                #region API. VISITORS ------------------------------------------------------------------------------
                if (dataType == "VISITORS")
                    {
                        //RESULT = dataType;
                        aceVisitor = new ACEVisitors(ace);
                        for (int i = 0; i < dtOutputData.Rows.Count; i++)
                        {
                            value = dtOutputData.Rows[i].ItemArray[2].ToString();
                            if (!String.IsNullOrEmpty(value) && value != "NULL")
                                switch (i)
                                {                                                                                                   // table [PERSONS]  - begin
                                    case 0: aceID = value;
                                        if (aceVisitor.Get(aceID) == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) F_API_ONLINE = false; 
                                        if (aceVisitor.Get(aceID) == API_RETURN_CODES_CS.API_SUCCESS_CS) F_RECORD_EXISTS = true; 
                                        break;
                                    case 1: aceVisitor.LASTNAME = value; break;
                                    case 2: aceVisitor.FIRSTNAME = value; break;
                                    case 3: aceVisitor.ADDITIONALLASTNAME = value; break;
                                    case 4: aceDateOfBirth = new ACEDateT(); aceDateOfBirth.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); aceVisitor.DATEOFBIRTH = aceDateOfBirth; break; //"yyyyMMdd"
                                    case 5: aceVisitor.PERSNO = value; break;
                                    case 6: aceVisitor.SEX = (ACESexT)Convert.ToInt32(value); break;
                                    case 7: aceVisitor.GRADE = value; break;
                                    case 8: aceVisitor.NUMBERPLATE = value; break;
                                    case 9: aceVisitor.STREETHOUSENO = value; break;
                                    case 10: aceVisitor.ZIPCODE = value; break;
                                    case 11: aceVisitor.CITY = value; break;
                                    case 12: aceVisitor.COUNTRY = value; break;
                                    case 13: aceVisitor.NATIONALITY = value; break;
                                    case 14: aceVisitor.PHONEPRIVATE = value; break;
                                    case 15: aceVisitor.FAXPRIVATE = value; break;
                                    case 16: aceVisitor.PHONEOFFICE = value; break;
                                    case 17: aceVisitor.FAXOFFICE = value; break;
                                    case 18: aceVisitor.PHONEMOBILE = value; break;
                                    case 19: aceVisitor.PHONEOTHER = value; break;
                                    case 20: aceVisitor.EMAIL = value; break;
                                    case 21: aceVisitor.WEBPAGEURL = value; break;
                                    case 22: aceVisitor.MAIDENNAME = value; break;
                                    case 23: aceVisitor.CITYOFBIRTH = value; break;
                                    case 24: aceVisitor.MARITALSTATUS = (ACEMaritalStateT)Convert.ToInt32(value); break;
                                    case 25: aceVisitor.IDTYPE = value; break;
                                    case 26: aceVisitor.IDNUMBER = value; break;
                                    case 27: aceIdValidUntil = new ACEDateT(); aceIdValidUntil.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); aceVisitor.IDVALIDUNTIL = aceIdValidUntil; break; //"yyyyMMdd"
                                    case 28: aceVisitor.HEIGHT = Convert.ToInt32(value); break;
                                    case 29: aceVisitor.DEPARTMENT = value; break;
                                    case 30: aceVisitor.CENTRALOFFICE = value; break;
                                    case 31: aceVisitor.COSTCENTRE = value; break;
                                    case 32: aceVisitor.JOB = value; break;
                                    case 33: aceVisitor.ATTENDANT = value; break;
                                    case 34: aceVisitor.REASONSTAY = value; break;
                                    case 35: aceVisitor.REMARK = value; break;
                                    case 36: aceVisitor.RESERVE1 = value; break;
                                    case 37: aceVisitor.RESERVE2 = value; break;
                                    case 38: aceVisitor.RESERVE3 = value; break;
                                    case 39: aceVisitor.RESERVE4 = value; break;
                                    case 40: aceVisitor.RESERVE5 = value; break;
                                    case 41: aceVisitor.RESERVE6 = value; break;
                                    case 42: aceVisitor.RESERVE7 = value; break;
                                    case 43: aceVisitor.RESERVE8 = value; break;
                                    case 44: aceVisitor.RESERVE9 = value; break;
                                    case 45: aceVisitor.RESERVE10 = value; break;                                         // table [PERSONS]  - end
                                    case 46: aceVisitor.IDENTIFIEDBY = (ACEVisitorIdentifyT)Convert.ToInt32(value); break;   // table [VISITORS] - begin
                                    case 47: aceVisitor.IDENTPAPER = value; break;
                                    case 48: aceVisitor.PASSPORTNO = value; break;
                                    case 49: aceVisitor.REASON = value; break;
                                    case 50: aceVisitor.LOCATION = value; break;
                                    case 51: aceVisitor.REMARKS = value; break;
                                    case 52: break;                         //CARDID
                                    case 53: DTime = new DateTime(); DTime = DateTime.ParseExact(value, "yyyyMMdd", null); aceVisitor.AUTHFROM = DTime; break; //"yyyyMMdd"
                                    case 54: DTime = new DateTime(); DTime = DateTime.ParseExact(value, "yyyyMMdd", null); aceVisitor.AUTHUNTIL = DTime; break; //"yyyyMMdd"
                                    case 55: ACEDateT ARRIVALEXPECTED = new ACEDateT(); ARRIVALEXPECTED.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); aceVisitor.ARRIVALEXPECTED = ARRIVALEXPECTED; break; //"yyyyMMdd"
                                    case 56: ACEDateT DEPARTEXPECTED = new ACEDateT(); DEPARTEXPECTED.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); aceVisitor.DEPARTEXPECTED = DEPARTEXPECTED; break; //"yyyyMMdd"
                                    case 57: aceVisitor.COMPANYID = value; break; //41_COMPANYID
                                    default: break;                                                                                  // table [VISITORS] - end
                                }
                        }
                    }

                #endregion-------------------------------------------------------------------------------------------------------------
                #region API. AUTHORIZATIONS ------------------------------------------------------------------------------
                if (dataType == "AUTHORIZATIONS")
                    {
                        //RESULT = dataType;
                        aceAuthorization = new ACEAuthorizations(ace);
                        for (int i = 0; i < dtOutputData.Rows.Count; i++)
                        {
                            value = dtOutputData.Rows[i].ItemArray[2].ToString();
                            if (!String.IsNullOrEmpty(value) && value != "NULL")
                                switch (i)
                                {
                                    case 0: aceID = value;                                                                                  //AUTHID
                                        if (aceAuthorization.Get(aceID) == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) F_API_ONLINE = false; 
                                        if (aceAuthorization.Get(aceID) == API_RETURN_CODES_CS.API_SUCCESS_CS) F_RECORD_EXISTS = true; 
                                        break; 
                                    case 1: aceAuthorization.SHORTNAME = value; break;
                                    case 2: aceAuthorization.NAME = value; break;
                                    case 3: aceAuthorization.TMID = value; break;
                                    case 4: aceAuthorization.CLIENTID = value; break;
                                    case 5: aceAuthorization.SPECIALFUNCTIONID = value; break;
                                    case 6: aceAuthorization.MACID = value;
                                        var query = new ACEQuery(ace);                                                          // --- GET MACID BEFORE ADDING AUTHORIZATION!
                                        query.Select("id", "devices", "DATEDELETED is NULL AND type=’MAC’");
                                        while (query.FetchNextRow()) { ACEColumnValue MACID = null; query.GetRowData(0, MACID); } // --- OK, MACID is ready, go on
                                        break;
                                    default: break;
                                }
                        }
                    }

                #endregion---------------
                #region API. CARDS ------------------------------------------------------------------------------
                if (dataType == "CARDS")
                    {
                        //RESULT = dataType;
                        aceCard = new ACECards(ace);
                        aceCard.CODEDATA2 = null;
                        aceCard.CODEDATA3 = null;
                        aceCard.CODEDATA4 = null;
                        F_VISITOR = false;
                        F_PERSON = false;
                        for (int i = 0; i < dtOutputData.Rows.Count; i++)
                        {
                            value = dtOutputData.Rows[i].ItemArray[2].ToString();
                            if (!String.IsNullOrEmpty(value) && value != "NULL")
                                switch (i)
                                {
                                    case 0: aceID = value;                                                                                      //CARDID
                                        if (aceCard.Get(aceID) == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) F_API_ONLINE = false; 
                                        if (aceCard.Get(aceID) == API_RETURN_CODES_CS.API_SUCCESS_CS) F_RECORD_EXISTS = 
                                            true; break;                  
                                    case 1: aceCard.CARDNO = value; break;                                                                 //CARDNO
                                    case 2: aceCard.CODEDATA = value; break;                                                                 //CODEDATA
                                    case 3: if (!String.IsNullOrEmpty(value) && value != "NULL") F_VISITOR = true; aceVisitor = new ACEVisitors(ace); aceVisitor.Get(value); break;   //41_VISID.  If VISID exists then this is visitor
                                    case 4: if (!String.IsNullOrEmpty(value) && value != "NULL") F_PERSON = true; acePerson = new ACEPersons(ace); acePerson.Get(value); break;   //41_PERSID. If PERSID or VISID exists then this is assigned card.
                                    default: break;
                                }
                        }
                    }

                #endregion-----------------------
                #region API. AUTHPERPERSON ------------------------------------------------------------------------------
                if (dataType == "AUTHPERPERSON")
                {
dp.logger("cp04: API_SendData_Single: API. region API. AUTHPERPERSON", "11000100"); 
                    //RESULT = "AUTHORIZATIONS ASSIGNMENT";
                    acePerson = new ACEPersons(ace);
                    aceAuthorization = new ACEAuthorizations(ace);
                    aceDateFrom = new ACEDateT();
                    aceDateTill = new ACEDateT();
                    for (int i = 0; i < dtOutputData.Rows.Count; i++)
                    {
                        value = dtOutputData.Rows[i].ItemArray[2].ToString();
                        if (!String.IsNullOrEmpty(value) && value != "NULL")
                            switch (i)
                            {
                                case 0: acePersonID = value; break;                             //41_PERSID
                                case 1: aceAuthID = value; break;                             //41_AUTHID
                                case 2: aceDateFrom.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); break;  //"dd.MM.yyyy HH:mm:ss"
                                case 3: aceDateTill.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); break;  //"dd.MM.yyyy HH:mm:ss"
                                default: break;
                            }
                    }
                }
                #endregion-------------------------------------------------------------------------------------------
                #region API. AUTHPERVISITOR ------------------------------------------------------------------------------
                if (dataType == "AUTHPERVISITOR")
                {
                    //RESULT = "AUTHORIZATIONS ASSIGNMENT";
                    aceVisitor = new ACEVisitors(ace);
                    aceAuthorization = new ACEAuthorizations(ace);
                    aceDateFrom = new ACEDateT();
                    aceDateTill = new ACEDateT();
                    for (int i = 0; i < dtOutputData.Rows.Count; i++)
                    {
                        value = dtOutputData.Rows[i].ItemArray[2].ToString();
                        if (!String.IsNullOrEmpty(value) && value != "NULL")
                            switch (i)
                            {
                                case 0: aceVisitorID = value; break;                             //41_VISID
                                case 1: aceAuthID = value; break;                             //41_AUTHID
                                case 2: aceDateFrom.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); break; //"dd.MM.yyyy HH:mm:ss"
                                case 3: aceDateTill.Set(dp.DAY(value, ""), dp.MONTH(value, ""), dp.YEAR(value, "")); break; //"dd.MM.yyyy HH:mm:ss"
                                default: break;
                            }
                    }
                }
                #endregion----------------------------------------------------------------------------------------------
                #region API. SEND DATA ---------------------------------------------------------------------------
                
                API_RETURN_CODES_CS result = API_RETURN_CODES_CS.API_CMD_FAILED_CS;

                if (cb_TestRun.Checked)
                {
                    RESULT = RESULT + "; TESTRUN; TEST; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                    count_PROCESSED++;
                    F_PROCESSED = true;
                }

                if (!cb_TestRun.Checked && !F_API_ONLINE) { REC_N--; return true; }                

                if (!cb_TestRun.Checked && F_API_ONLINE && !F_RECORD_EXISTS && cb_Add.Checked)
                {
                    if (dataType == "AUTHORIZATIONS")   result = aceAuthorization.Add();
                    if (dataType == "EMPLOYEES")        result = acePerson.Add();
                    if (dataType == "VISITORS")         result = aceVisitor.Add();
                    if (dataType == "COMPANIES")        result = aceCompany.Add();
                    if (dataType == "TIME MODELS")      result = aceTimeModel.Add();
                    if (dataType == "CARDS")
                    {
                        if (!F_VISITOR && !F_PERSON)    result = aceCard.Add();                                                                             // card unassigned
                        if (F_PERSON)
                        {
                            aceCard.Add();
                            aceCardID = dp.dtFindDB(TB_DBCONNFILE_PROCESSOR.Text, TB_WORKFILESPATH.Text + @"\_sqlQueries\" + "bis41.extRQ.06.sql", aceCard.CARDNO.Trim(), aceCard.CODEDATA.Trim(), "", ""); //Get 41_CARDID of the last added card
                            result = aceCard.Get(aceCardID);
                            if (result == API_RETURN_CODES_CS.API_SUCCESS_CS && F_VISITOR) result = aceVisitor.AddCard(aceCard.GetCardId());   // card assigned to visitor 
                            if (result == API_RETURN_CODES_CS.API_SUCCESS_CS && !F_VISITOR) result = acePerson.AddCard(aceCard.GetCardId());    // card assigned to employee 
                            //DataRow rowEdit = dtOutputData.Rows[0]; rowEdit.BeginEdit(); rowEdit["DATA"] = aceCardID; rowEdit.EndEdit();      // change displayed CARDID from 23_CARDID to 41_CARDID (debugging function, delete later)
                        }
                    }
                    if (dataType == "AUTHPERPERSON")
                    {
                        //if (acePerson.Get(acePersonID) == API_RETURN_CODES_CS.API_SUCCESS_CS && aceAuthorization.Get(aceAuthID) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                        //    result = acePerson.AddAuthorization(aceAuthorization.GetAuthorizationId(), aceDateFrom, aceDateTill);                   //result = acePerson.AddAuthorization(aceAuthorization.GetAuthorizationId());
                        acePerson.Get(acePersonID);
                        result = acePerson.AddAuthorization(aceAuthID, aceDateFrom, aceDateTill);                   //result = acePerson.AddAuthorization(aceAuthorization.GetAuthorizationId());
                    }
                    if (dataType == "AUTHPERVISITOR")
                    {
                        //if (aceVisitor.Get(aceVisitorID) == API_RETURN_CODES_CS.API_SUCCESS_CS && aceAuthorization.Get(aceAuthID) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                        //    result = aceVisitor.AddAuthorization(aceAuthorization.GetAuthorizationId(), aceDateFrom, aceDateTill);                  //result = aceVisitor.AddAuthorization(aceAuthorization.GetAuthorizationId());
                        aceVisitor.Get(aceVisitorID);
                        result = aceVisitor.AddAuthorization(aceAuthID, aceDateFrom, aceDateTill);                  //result = aceVisitor.AddAuthorization(aceAuthorization.GetAuthorizationId());
                    }
                    RESULT = RESULT + "; ADD   ; " + result.ToString() + "; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                    if (result == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) { F_API_ONLINE = false; REC_N--; return true; } 
                    if (result != API_RETURN_CODES_CS.API_SUCCESS_CS) count_ERRORS++;
                    count_ADDED++;
                    count_PROCESSED++;
                    F_PROCESSED = true;
                }

                if (!cb_TestRun.Checked && F_API_ONLINE && F_RECORD_EXISTS && cb_Skip.Checked)
                    {
                        RESULT = RESULT + "; SKIPPED; " + result.ToString() + "; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                        count_SKIPPED++;
                        count_PROCESSED++;
                        F_PROCESSED = true;
                    }

                if (!cb_TestRun.Checked && F_API_ONLINE && F_RECORD_EXISTS && cb_Update.Checked && !cb_Skip.Checked)
                    {
                        if (dataType == "AUTHORIZATIONS")   result = aceAuthorization.Update();
                        if (dataType == "EMPLOYEES")        result = acePerson.Update();
                        if (dataType == "VISITORS")         result = aceVisitor.Update();
                        if (dataType == "COMPANIES")        result = aceCompany.Update();
                        if (dataType == "TIME MODELS")      result = aceTimeModel.Update();
                        if (dataType == "CARDS")            result = aceCard.Update();
                        if (dataType == "AUTHPERPERSON")    result = API_RETURN_CODES_CS.API_CMD_FAILED_CS; // no logic for 'Update'
                        if (dataType == "AUTHPERVISITOR")   result = API_RETURN_CODES_CS.API_CMD_FAILED_CS; // no logic for 'Update'
                        RESULT = RESULT + "; UPDATE; " + result.ToString() + "; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                        if (result == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) { F_API_ONLINE = false; REC_N--; return true; } 
                        if (result != API_RETURN_CODES_CS.API_SUCCESS_CS) count_ERRORS++;    
                        count_UPDATED++;
                        count_PROCESSED++;
                        F_PROCESSED = true;
                    }

                if (!cb_TestRun.Checked && F_API_ONLINE && F_RECORD_EXISTS && cb_Delete.Checked && !cb_Skip.Checked)
                    {
                        if (dataType == "AUTHORIZATIONS")   result = aceAuthorization.Delete();
                        if (dataType == "EMPLOYEES")        result = acePerson.Delete();
                        if (dataType == "VISITORS")         result = aceVisitor.Delete();
                        if (dataType == "COMPANIES")        result = aceCompany.Delete();
                        if (dataType == "TIME MODELS")      result = aceTimeModel.Delete();
                        if (dataType == "CARDS")            result = aceCard.Delete();
                        if (dataType == "AUTHPERPERSON")
                        {
                            if (acePerson.Get(acePersonID) == API_RETURN_CODES_CS.API_SUCCESS_CS && aceAuthorization.Get(aceAuthID) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                                result = acePerson.RemoveAuthorization(aceAuthorization.GetAuthorizationId());
                        }
                        if (dataType == "AUTHPERVISITOR")
                        {
                            if (aceVisitor.Get(aceVisitorID) == API_RETURN_CODES_CS.API_SUCCESS_CS && aceAuthorization.Get(aceAuthID) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                                result = aceVisitor.RemoveAuthorization(aceAuthorization.GetAuthorizationId());
                        }
                        RESULT = RESULT + "; DELETE; " + result.ToString() + "; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                        if (result == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) { F_API_ONLINE = false; REC_N--; return true; }  
                        if (result != API_RETURN_CODES_CS.API_SUCCESS_CS) count_ERRORS++;    
                        count_DELETED++;
                        count_PROCESSED++;
                        F_PROCESSED = true;
                    }

                if (!F_PROCESSED)
                {
                    RESULT = RESULT + "; SKIPPED; " + "Nothing to do" + "; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                    count_SKIPPED++;
                    count_PROCESSED++;
                    F_PROCESSED = false;
                }

                if (result == API_RETURN_CODES_CS.API_LOGIN_REQUIRED_CS) { F_API_ONLINE = false; REC_N--; return true; }

                lbl_API_Result.Text     = RESULT;
                TB_ADDED.Text           = count_ADDED.ToString();
                TB_UPDATED.Text         = count_UPDATED.ToString();
                TB_DELETED.Text         = count_DELETED.ToString();
                TB_SKIPPED.Text         = count_SKIPPED.ToString();
                TB_PROCESSED.Text       = count_PROCESSED.ToString();
                TB_ERRORS.Text          = count_ERRORS.ToString();

                if (cb_LogRAM.Checked)  dp.logger(RESULT, "11000010");
                else                    global.Data.logger(RESULT, "10000010");
                return true;
                #endregion---------------------
                }
                catch (Exception e)
                {
                    msgError = "API_SendData_Single. \r\n" +
                                dataType + "\r\n" +
                                e.Message + "\r\n" +
                                "1. Check if output file was converted properly\r\n2. Check correspodnence between fields of output file and schema file" + "\r\n" +
                                "Options: Abort the application or Retry to start new operation or Ignore this message";
                    msgErrorType = "GENERAL EXCEPTION";
                    global.Data.logger(msgError, "10000100");
                    DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    if (DialogResult == DialogResult.Abort)     { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                    if (DialogResult == DialogResult.Retry)     { global.Data.logger(DialogResult.ToString(), "10000100"); return false; }
                    if (DialogResult == DialogResult.Ignore)    { global.Data.logger(DialogResult.ToString(), "10000100"); return true; }
                }
            return false;
        }
        private bool SQL_SendData_Single(DataTable dtOutputFile, DataTable dtAPISchemaFile, string dataType, int REC_N)
        {
            #region SQL. Define common variables ------------------------------------------------------------------------

            string RESULT = REC_N.ToString() + "; " + dataType;
            string  SQLRESULT   = "";
            bool    TestRun     = cb_TestRun.Checked;
            bool    Add         = cb_Add.Checked;
            bool    F_RECORD_EXISTS      = false;
            string  connString  = "";
            string  sqlRq       = "";
            string Data0, Data1, Data2, Data3, Data4, Data5, Data6;
            Data0 = Data1 = Data2 = Data3 = Data4 = Data5 = Data6 = "";

            #endregion--------------

            //if (!displayOutputData(dtOutputFile, dtAPISchemaFile, REC_N)) return false;          //Prepare and display output data

            connString  = dp.txtReadFromFile(TB_DBCONNFILE_OUT.Text);
            sqlRq       = dp.txtReadFromFile(TB_SQLFILE_OUT.Text);

            if (displayOutputData(dtOutputFile, dtAPISchemaFile, REC_N) && !String.IsNullOrEmpty(connString) && !String.IsNullOrEmpty(sqlRq))  //Prepare and display output data. If output data is ready then continue (assign it to api variables)
            try
            {
            using (SqlConnection dbConn = new SqlConnection(connString))
                {
                    dbConn.Open();
                    SqlCommand SQLcmd = new SqlCommand(sqlRq, dbConn);
                    #region SQL. RCPPERAUTH-------------------------------------------------------------------
                    if (dataType == "RCPPERAUTH")
                    {
                        Data0 = dtOutputData.Rows[0].ItemArray[2].ToString();
                        Data1 = dtOutputData.Rows[1].ItemArray[2].ToString();
                        Data2 = dtOutputData.Rows[2].ItemArray[2].ToString();
                        SqlParameter data0 = new SqlParameter("@Data0", Data0);
                        SqlParameter data1 = new SqlParameter("@Data1", Data1);
                        SqlParameter data2 = new SqlParameter("@Data2", Data2);
                        SQLcmd.Parameters.Add(data0);
                        SQLcmd.Parameters.Add(data1);
                        SQLcmd.Parameters.Add(data2);
                    }
                    #endregion--------------
                    #region SQL. CARDS-------------------------------------------------------------------
                    if (dataType == "CARDS")
                    {
                        Data0 = dtOutputData.Rows[0].ItemArray[2].ToString();
                        Data1 = dtOutputData.Rows[1].ItemArray[2].ToString();
                        Data2 = dtOutputData.Rows[2].ItemArray[2].ToString();
                        Data3 = dtOutputData.Rows[3].ItemArray[2].ToString();
                        Data4 = dtOutputData.Rows[4].ItemArray[2].ToString();
                        Data5 = dtOutputData.Rows[5].ItemArray[2].ToString();
                        Data6 = dtOutputData.Rows[6].ItemArray[2].ToString();
                        SqlParameter data0 = new SqlParameter("@Data0", Data0);
                        SqlParameter data1 = new SqlParameter("@Data1", Data1);
                        SqlParameter data2 = new SqlParameter("@Data2", Data2);
                        SqlParameter data3 = new SqlParameter("@Data3", Data3);
                        SqlParameter data4 = new SqlParameter("@Data4", Data4);
                        SqlParameter data5 = new SqlParameter("@Data5", Data5);
                        SqlParameter data6 = new SqlParameter("@Data6", Data6);
                        SQLcmd.Parameters.Add(data0);
                        SQLcmd.Parameters.Add(data1);
                        SQLcmd.Parameters.Add(data2);
                        SQLcmd.Parameters.Add(data3);
                        SQLcmd.Parameters.Add(data4);
                        SQLcmd.Parameters.Add(data5);
                        SQLcmd.Parameters.Add(data6);
                    }
                    #endregion--------------
                    #region SQL. SEND DATA ------------------------------------------------------------------------------
                    if (cb_TestRun.Checked)
                    {
                        RESULT = RESULT + "; TESTRUN; TEST; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                        count_PROCESSED++;
                        F_PROCESSED = true;
                    }

                    if (!cb_TestRun.Checked && F_RECORD_EXISTS && cb_Skip.Checked)
                    {
                        RESULT = RESULT + "; SKIPPED; *****; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                        count_SKIPPED++;
                        count_PROCESSED++;
                        F_PROCESSED = true;
                    }

                    if (!cb_TestRun.Checked && !F_RECORD_EXISTS && cb_Add.Checked)
                    {
                        if (dataType == "RCPPERAUTH") 
                            try
                            {
                                SQLRESULT = SQLcmd.ExecuteNonQuery().ToString();
                                RESULT = RESULT + "; SQL DATA" + "; ADD; " + SQLRESULT + "; " + Data0 + "; " + Data1 + "; " + Data2;
                            }
                            catch (SqlException se)
                            {
                                RESULT = RESULT + "; SQL DATA" + "; ADD; ERROR;" + Data0 + "; " + Data1 + "; " + Data2 + "; " + se.Message;
                                count_ERRORS++;
                            }
                        if (dataType == "CARDS") 
                            try
                            {
                                SQLRESULT = SQLcmd.ExecuteNonQuery().ToString();
                                RESULT = RESULT + "; SQL DATA" + "; ADD; " + SQLRESULT + "; " + Data0 + "; " + Data1 + "; " + Data2;
                            }
                            catch (SqlException se)
                            {
                                RESULT = RESULT + "; SQL DATA" + "; ADD; ERROR;" + Data0 + "; " + Data1 + "; " + Data2 + "; " + se.Message;
                                count_ERRORS++;
                            }
                        count_ADDED++;
                        count_PROCESSED++;
                        F_PROCESSED = true;
                    }
                    dbConn.Close();
                }
            lbl_API_Result.Text         = RESULT;
            TB_ADDED.Text               = count_ADDED.ToString();
            TB_UPDATED.Text             = count_UPDATED.ToString();
            TB_DELETED.Text             = count_DELETED.ToString();
            TB_SKIPPED.Text             = count_SKIPPED.ToString();
            TB_PROCESSED.Text           = count_PROCESSED.ToString();
            TB_ERRORS.Text              = count_ERRORS.ToString();

            //if (cb_LogRAM.Checked) dp.logger(RESULT, "11000010");
            //else 
                global.Data.logger(RESULT, "10000010");
            return true;
            #endregion-------------------
            }
            catch (SqlException se)
            {
                msgError = "SQL_SendData_Single. \r\n" +
                            dataType + "\r\n" +
                            se.Message + "\r\n" +
                            "1. Check if output file was converted properly\r\n2. Check correspodnence between fields of output file and schema file\r\n3. Check accessibility of database" + "\r\n" +
                            "Options: Abort the application or Retry to start new operation or Ignore this message";
                msgErrorType = "GENERAL EXCEPTION";
                global.Data.logger(msgError, "10000100");
                DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); return false; }
                if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); return true; }
            }
            return false;
        }
        private bool CS_SendData_Single(DataTable dtOutputFile, DataTable dtAPISchemaFile, string dataType, int REC_N)
        {
            #region CS. Define common variables ------------------------------------------------------------------------
            
            bool F_RECORD_EXISTS    = false;
            string RESULT           = REC_N.ToString() + "; " + dataType;
            string                  Data1, Data2, Data3;
            Data1 = Data2 = Data3 = "";

            #endregion--------------

            if (displayOutputData(dtOutputFile, dtAPISchemaFile, REC_N))                            //Prepare and display output data. If output data is ready then continue (assign it to api variables)
            try
            {
            #region CS. PHOTOS-------------------------------------------------------------------

                if (dataType == "PHOTOS")
                {
                    Data1 = TB_PHOTOS_IN.Text + @"\" + dtOutputData.Rows[0].ItemArray[2].ToString() + ".jpg";       // Source file
                    Data2 = TB_PHOTOS_OUT.Text + @"\" + dtOutputData.Rows[1].ItemArray[2].ToString() + ".jpg";      // Destination file
                }
            #endregion--------------
            #region CS. SEND DATA ------------------------------------------------------------------------------
                
                if (cb_TestRun.Checked)
                {
                    RESULT = RESULT + "; TESTRUN; TEST; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                    count_PROCESSED++;
                    F_PROCESSED = true;
                }

                if (!cb_TestRun.Checked && F_RECORD_EXISTS && cb_Skip.Checked)
                {
                    RESULT = RESULT + "; SKIPPED; *****; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                    count_SKIPPED++;
                    count_PROCESSED++;
                    F_PROCESSED = true;
                }
                
                if (!cb_TestRun.Checked && !F_RECORD_EXISTS && cb_Add.Checked)
                {
                    if (dataType == "PHOTOS")
                        try
                        {
                            FileInfo file = new FileInfo(Data1);
                            file.CopyTo(Data2, true);
                            RESULT = RESULT + "; CS; ADD; COPIED; " + Data1 + "; " + Data2 + "; " + Data3;
                        }
                        catch (IOException ioex)
                        {
                            string ioexERROR = ioex.Message;
                            RESULT = RESULT + "; CS; ADD; ERROR;  " + ioexERROR;
                            count_ERRORS++;
                        }
                    // For further programming: If any more 'ADD' operations exist put them here
                    
                    count_ADDED++;
                    count_PROCESSED++;
                    F_PROCESSED = true;
                }

                if (!F_PROCESSED)
                {
                    RESULT = RESULT + "; SKIPPED; " + "Nothing to do" + "; " + "ID: " + dtOutputData.Rows[0].ItemArray[2].ToString() + "; " + dtOutputData.Rows[1].ItemArray[2].ToString() + "; " + dtOutputData.Rows[2].ItemArray[2].ToString();
                    count_SKIPPED++;
                    count_PROCESSED++;
                    F_PROCESSED = false;
                }

                lbl_API_Result.Text     = RESULT;
                TB_ADDED.Text           = count_ADDED.ToString();
                TB_UPDATED.Text         = count_UPDATED.ToString();
                TB_DELETED.Text         = count_DELETED.ToString();
                TB_SKIPPED.Text         = count_SKIPPED.ToString();
                TB_PROCESSED.Text       = count_PROCESSED.ToString();
                TB_ERRORS.Text          = count_ERRORS.ToString();

                //if (cb_LogRAM.Checked)  dp.logger(RESULT, "11000010");        // RAM logging: possibly can be programmed later
                //else                    
                    global.Data.logger(RESULT, "10000010");
                return true;
                #endregion---------------
            }
            catch (Exception e)
            {
                msgError = "CS_SendData_Single. \r\n" +
                            dataType + "\r\n" +
                            e.Message + "\r\n" +
                            "1. Check if output file was converted properly\r\n2. Check correspodnence between fields of output file and schema file" + "\r\n" +
                            "Options: Abort the application or Retry to start new operation or Ignore this message";
                msgErrorType = "GENERAL EXCEPTION";
                global.Data.logger(msgError, "10000100");
                DialogResult = MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                if (DialogResult == DialogResult.Abort) { global.Data.logger(DialogResult.ToString(), "10000100"); System.Diagnostics.Process.GetCurrentProcess().Kill(); }
                if (DialogResult == DialogResult.Retry) { global.Data.logger(DialogResult.ToString(), "10000100"); return false; }
                if (DialogResult == DialogResult.Ignore) { global.Data.logger(DialogResult.ToString(), "10000100"); return true; }
            }
            return false;
        }

#endregion ---------------------------------------------------------------------------------------

#region TEST DATA GENERATOR---------------------------------------------------------------------------------

        private void btn_GenerateObjects_Click(object sender, EventArgs e)
        {
            string saveToFile = TB_WORKFILESPATH.Text + @"\_dataIn\" + TB_TESTFILENAME.Text;
            dp.dtSaveToFile(dtGenerateObjects(Convert.ToInt32(TB_QTYOFOBJECTS.Text), TB_PREF.Text, TB_TYPE.Text, TB_END.Text), saveToFile, true);
        }
        public DataTable dtGenerateObjects(int qty, string pref, string type, string ending)
        {
            DataTable dtRead = new DataTable("dt");
            dtRead.Clear(); dtRead.Columns.Clear();
            string[,] strTable = new string[500, 200000];
            pb_ObjectGenerator.Maximum = qty;

            //string Line;
            int lastColumn = 50;
            string[] strArrIn = new string[500];

            if (true) try
                {
                    DataColumn[] strArrColumnsIn = new DataColumn[lastColumn + 1];
                    for (int i = 0; i <= lastColumn; i++)
                    {
                        strArrColumnsIn[i] = new DataColumn(i.ToString(), typeof(String));
                        dtRead.Columns.Add(strArrColumnsIn[i]);
                    }
                    DataRow dr = null;
                    for (int j = 1; j <= qty; j++)
                    {
                        dr = dtRead.NewRow();
                        for (int i = 0; i <= lastColumn; i++)
                        {
                            switch (i)
                            {
                                case 0: dr[i] = j.ToString(); break;
                                case 1: dr[i] = "LASTNAME " + pref + type + ending + j.ToString(); break;
                                case 2: dr[i] = "FIRSTNAME " + type + ending + j.ToString(); break;
                                case 3: dr[i] = "ADDITIONALLASTNAME " + type + ending + j.ToString(); break;
                                case 4: dr[i] = "19730307"; break;
                                case 5: dr[i] = j.ToString(); break;
                                case 6: dr[i] = "0"; break;
                                case 7: dr[i] = "Mr. " + type + ending + j.ToString(); break;
                                case 8: dr[i] = j.ToString(); break;
                                case 9: dr[i] = j.ToString(); break;
                                case 10: dr[i] = j.ToString(); break;
                                case 11: dr[i] = "CITY " + type + ending + j.ToString(); break;
                                case 12: dr[i] = "COUNTRY " + type + ending + j.ToString(); break;
                                case 13: dr[i] = "NATIONALITY " + type + ending + j.ToString(); break;
                                case 14: dr[i] = j.ToString(); break;
                                case 15: dr[i] = j.ToString(); break;
                                case 16: dr[i] = j.ToString(); break;
                                case 17: dr[i] = j.ToString(); break;
                                case 18: dr[i] = j.ToString(); break;
                                case 19: dr[i] = j.ToString(); break;
                                case 20: dr[i] = type + j.ToString() + @"@mail.ru"; break;
                                case 21: dr[i] = "www." + pref + type + j.ToString() + ".ru"; break;
                                case 24: dr[i] = "1"; break;
                                case 25: dr[i] = "1"; break;
                                case 27: dr[i] = "20170307"; break;
                                case 28: dr[i] = "175"; break;
                                default: dr[i] = j.ToString(); break;
                            }
                        }
                        dtRead.Rows.Add(dr);
                        pb_ObjectGenerator.Value = j;
                    }
                    return dtRead;
                }
                catch (Exception e) { MessageBox.Show("dtReadFromFile \nFill DataTable: \n" + e.Message, "GENERAL EXCEPTION", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
            return dtRead;
        }

#endregion ------------------------------------------------------------------------------

#region APPENDIX ------------------------------------------------------------------------------

        private void btn_ResetStatistics_Click(object sender, EventArgs e)
        {
            ResetStatistics();
        }
        private void ResetStatistics()
        {
            count_PROCESSED     =
            count_ADDED         =
            count_SKIPPED       =
            count_UPDATED       =
            count_DELETED       = 
            count_ERRORS        = 0;
            TB_PROCESSED.Text   =
            TB_ADDED.Text       =
            TB_SKIPPED.Text     =
            TB_UPDATED.Text     =
            TB_DELETED.Text     =
            TB_ERRORS.Text      = "0";
        }
        private void TB_LOGAPI_TextChanged(object sender, EventArgs e)
        {
            global.Data.LOGAPI = this.TB_LOG_API.Text;
        }
        private void TB_LOGAPP_TextChanged(object sender, EventArgs e)
        {
            global.Data.LOGAPP = this.TB_LOGAPP.Text;
        }
        private void TB_WORKFILESPATH_TextChanged(object sender, EventArgs e)
        {
            global.Data.WORKFILESPATH = this.TB_WORKFILESPATH.Text;
        }
        private void TB_INPUTFILE_TextChanged(object sender, EventArgs e)
        {
            //global.Data.INPUTFILE = this.TB_INPUTFILE.Text;
        }
        //private void button2_Click(object sender, EventArgs e)
        //{

        //    DataTable dt_asf = new DataTable("dtasf");
        //    dt_asf.Clear(); dt_asf.Columns.Clear();
        //    dt_asf = dp.dtReadFromFile(@"\\Esb\host. documents\VTB\_associations\scm1.authorizations.asf.csv");

        //    //DataRow myRowShow = dt_asf.Rows[2];
        //    int show = Convert.ToInt32(dt_asf.Rows[15].ItemArray[1].ToString());
        //    MessageBox.Show(show.ToString());
        //}
        private void btnAPILogin2_Click(object sender, EventArgs e)
        {
            APILogin(TB_API_LOGIN_SERVER.Text, API_LOGIN_NAME.Text, API_LOGIN_PWD.Text);
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cbInputSource.SelectedItem.ToString())
            {
                case "SQL DB":
                    //this.TB_INPUTFILE.ReadOnly = true;
                    //this.TB_DBCONNFILEOPEN.ReadOnly = false;
                    break;
                case "Input File":
                    //this.TB_INPUTFILE.ReadOnly = false;
                    //this.TB_DBCONNFILEOPEN.ReadOnly = true;

                    break;
                default:
                    this.lbl_APP_Result.Text = "Job: No possible data source found!";
                    break;
            }
        }
        private void btnOpenAppLog_Click(object sender, EventArgs e)
        {
            //          Process.Start(@"C:\Windows\notepad.exe " + this.TB_WORKFILESPATH.Text + @"\_logs\" + this.TB_LOGAPP.Text);
        }
        private void btnOpenAPILog_Click(object sender, EventArgs e)
        {
            //            Process.Start(@"C:\Windows\notepad.exe " + this.TB_WORKFILESPATH.Text + @"\_logs\" + this.TB_LOGAPI.Text);
        }
        private void checkBoxLogAPIWrite_CheckedChanged(object sender, EventArgs e)
        {
            global.Data.F_LOGAPI = cb_LogAPI.Checked; // possibly obsolette
            initiateLogging();
        }
        private void checkBoxLogDBWrite_CheckedChanged(object sender, EventArgs e)
        {
            global.Data.F_LOGAPP = cb_LogAPP.Checked; // possibly obsolette
            initiateLogging();
        }
        private void cb_AutoRun_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_AutoRun.Checked || cb_Single.Checked)
            {
                this.TB_API_PROCESSOR_STARTLINE.ReadOnly = true;
                this.TB_API_PROCESSOR_LINES.ReadOnly = true;
            }
            else
            {
                this.TB_API_PROCESSOR_STARTLINE.ReadOnly = false;
                this.TB_API_PROCESSOR_LINES.ReadOnly = false;
            }

            if (cb_AutoRun.Checked && API_PROCESSOR_set_parameters())
            {
                cb_Single.Checked = false;
                btn_Start.Enabled = true;
                //highlight current cell in dgv
                //dgvOutputFile.Rows[Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text)].Cells[0].Selected = true;
                //dgvOutputFile.CurrentCell = dgvOutputFile.Rows[Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text)].Cells[0];
            }
            else
            {
                cb_AutoRun.Checked = false;
                btn_Start.Enabled = false;
            }
        }
        private void cb_Single_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_AutoRun.Checked || cb_Single.Checked)
            {
                this.TB_API_PROCESSOR_STARTLINE.ReadOnly = true;
                this.TB_API_PROCESSOR_LINES.ReadOnly = true;
            }
            else
            {
                this.TB_API_PROCESSOR_STARTLINE.ReadOnly = false;
                this.TB_API_PROCESSOR_LINES.ReadOnly = false;
            }

            if (cb_Single.Checked && API_PROCESSOR_set_parameters())
            {
                cb_AutoRun.Checked = false;
                btn_Start.Enabled = true;
                global.Data.REC_N = Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text);
                //highlight current cell in dgv
                //dgvOutputFile.Rows[Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text)].Cells[0].Selected = true;
                //dgvOutputFile.CurrentCell = dgvOutputFile.Rows[Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text)].Cells[0];
            }
            else
            {
                cb_Single.Checked = false;
                btn_Start.Enabled = false;
            }
        }
        private bool API_PROCESSOR_set_parameters()
        {
            int OUTPUTFILE_ROWS = 0;
            int APISCHEMAFILE_ROWS = 0;

            if (true) try
            {
                if (!String.IsNullOrEmpty(TB_OUTPUTFILE.Text)) OUTPUTFILE_ROWS = System.IO.File.ReadAllLines(TB_OUTPUTFILE.Text).Length;
                if (!String.IsNullOrEmpty(TB_APISCHEMAFILE.Text)) APISCHEMAFILE_ROWS = System.IO.File.ReadAllLines(TB_APISCHEMAFILE.Text).Length;
            }
            catch { MessageBox.Show("WRONG INPUT\n1. Check if output and schema files exist", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

            if (OUTPUTFILE_ROWS > 0 && APISCHEMAFILE_ROWS > 0) try
            {
                TB_OUTPUTFILE_ROWS.Text = OUTPUTFILE_ROWS.ToString();
                TB_APISCHEMAFILE_ROWS.Text = APISCHEMAFILE_ROWS.ToString();
                //TB_API_PROCESSOR_STARTLINE.Text = "0";
                //TB_API_PROCESSOR_LINES.Text     = (OUTPUTFILE_ROWS - 1).ToString();
                int API_PROCESSOR_STARTLINE = Convert.ToInt32(this.TB_API_PROCESSOR_STARTLINE.Text);
                int API_PROCESSOR_LINES = Convert.ToInt32(this.TB_API_PROCESSOR_LINES.Text);

                if (API_PROCESSOR_LINES == 0
                    || TB_API_PROCESSOR_LINES.Text == ""
                    || API_PROCESSOR_LINES > OUTPUTFILE_ROWS - 1
                    )
                    this.TB_API_PROCESSOR_LINES.Text = (OUTPUTFILE_ROWS - 1).ToString();

                if (API_PROCESSOR_STARTLINE + API_PROCESSOR_LINES > OUTPUTFILE_ROWS - 1) API_PROCESSOR_STARTLINE = 0;
                this.TB_API_PROCESSOR_STARTLINE.Text = API_PROCESSOR_STARTLINE.ToString();
                dp_updateProgressBarsMaximum(this, "SEND. API_PROCESSOR_set_parameters", -1, -1, Convert.ToInt32(this.TB_API_PROCESSOR_LINES.Text));
                dp_updateProgressBarsCurrent(this, "SEND. API_PROCESSOR_set_parameters ", -1, -1, 0);
                return true;
            }
            catch { MessageBox.Show("WRONG INPUT\n1. Check if output and schema files exist;\n2. Check if start line and number of lines are defined properly", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
            return false;
        }
        private void controlsStatus_Blocked()
        {
            //this.TB_WORKFILESPATH.ReadOnly = true;
            //this.TB_INPUTFILE.ReadOnly = true;
            //this.TB_OUTPUTFILE.ReadOnly = true;
            //this.TB_APIASSOCFILE.ReadOnly = true;
            //this.TB_APISCHEMAFILE.ReadOnly = true;
            //this.TB_SQLRQFILE.ReadOnly = true;
            //this.TB_DBCONNFILEOPEN.ReadOnly = true;
            //this.btnReadInputFile.Enabled = false;
            //this.btnRequestDB.Enabled = false;
            //this.btnSetInputFile.Enabled = false;
            //this.btnSetOutputFile.Enabled = false;
            //this.btnSetAPIAssocFile.Enabled = false;
            //this.btnSetAPISchemaFile.Enabled = false;
            //this.cbInputSource.Enabled = false;
            //this.cb_DataFormat.Enabled = false;
            //this.btn_API_Start.Enabled = false;
        }
        private void btn_JBE_Add_Click(object sender, EventArgs e)
        {
            string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JLE_JOBLIST.Text;
            dp.dgvSaveToFile(dgv_JOBLIST, jobListPath);
            loadJobList(jobListPath, true);
        }
        private void btn_JLE_Load_Click(object sender, EventArgs e)
        {
            DataTable dtJoblist = new DataTable("dtJoblist");
            dtJoblist.Clear(); dtJoblist.Columns.Clear();
            string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JLE_JOBLIST.Text;
            loadJobList(jobListPath, true);
            dtJoblist = dp.dtReadFromFile(jobListPath, true);
            dgv_JOBLIST.DataSource = dtJoblist;
        }
        private void dgv_JOBLIST_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (dgv_JOBLIST.CurrentRow != null) try
                {
                    int currentIndex = dgv_JOBLIST.CurrentCell.RowIndex;
                    //MessageBox.Show(currentIndex.ToString());

                    DataTable dtJoblist = new DataTable("dtJoblist");
                    dtJoblist.Clear(); dtJoblist.Columns.Clear();
                    string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JLE_JOBLIST.Text;
                    loadJobList(jobListPath, true);
                    dtJoblist = dp.dtReadFromFile(jobListPath, true);
                    dgv_JOBLIST.DataSource = dtJoblist;

                    DataRow dr = dtJoblist.Rows[currentIndex];
                    dtJoblist.ImportRow(dr);

                    dp.dgvSaveToFile(dgv_JOBLIST, jobListPath);
                    loadJobList(jobListPath, true);
                }
                catch (Exception ex)
                {
                    msgError = "dgv_JOBLIST. \r\n" +
                                "Clone Job" + "\r\n" +
                                ex.Message + "\r\n" +
                                "1. Check if joblist file name is properly indicated";
                    msgErrorType = "GENERAL EXCEPTION";
                    global.Data.logger(msgError, "10000100");
                    MessageBox.Show(msgError, msgErrorType, MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Exclamation);
                    if (this.DialogResult == DialogResult.Abort) { Application.Exit(); }
                }
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            if (dgv_JOBLIST.CurrentRow != null)
            {
                int currentIndex = dgv_JOBLIST.CurrentRow.Index;
                //MessageBox.Show(currentIndex.ToString());

                DataTable dtJoblist = new DataTable("dtJoblist");
                dtJoblist.Rows.Clear(); dtJoblist.Clear(); dtJoblist.Columns.Clear();
                string jobListPath = TB_WORKFILESPATH.Text + @"\_config\" + this.TB_JLE_JOBLIST.Text;
                loadJobList(jobListPath, true);
                dtJoblist = dp.dtReadFromFile(jobListPath, true);
                dgv_JOBLIST.DataSource = dtJoblist;

                DataRow dr = dtJoblist.Rows[currentIndex];
                dtJoblist.Rows.Remove(dr);

                dp.dgvSaveToFile(dgv_JOBLIST, jobListPath);
                loadJobList(jobListPath, true);
            }
        }
        private void button2_Click_2(object sender, EventArgs e)
        {
            //ACEEntranceInfo aceEntrance = new ACEEntranceInfo();
            //result = aceEntrance.m_strEntranceName(TB_ENTRANCEID.Text);
        }
        private void btn_ENTRANCE_UPDATE_Click(object sender, EventArgs e)
        {
            ACEEntranceInfo aceEntrance = new ACEEntranceInfo();
            //result = aceEntrance.m_strEntranceName();
            //.Get(TB_ENTRANCEID.Text);
        }
        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {
            //Add code here if hint pops up
        }
        private void interClass()
        {
            dp.DPName = "hello";
            dp.interClass();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            interClass();
        }
        private void dp_OnSomeNeeded(object sender, string inf, int bar1, int bar2, int bar3)
        {
            MessageBox.Show("Form: " + Name + "\n" + "Sender: " + sender.ToString() + "\n" + "Info: " + inf + "\n" + "bar1: " + bar1 + "\n" + "bar2: " + bar2 + "\n" + "bar3: " + bar3);
        }
        private void dp_updateProgressBarsCurrent(object sender, string inf, int bar1, int bar2, int bar3) // negative values do not change indications
        {
            if (bar1 <= pb_ReadData.Maximum && bar1 >= 0) pb_ReadData.Value = bar1;
            if (bar2 <= pb_ConvertData.Maximum && bar2 >= 0) pb_ConvertData.Value = bar2;
            if (bar3 <= pb_SendData.Maximum && bar3 >= 0) pb_SendData.Value = bar3;
            //MessageBox.Show("Form: " + Name + "\n" + "Sender: " + sender.ToString()+ "\n" + "CUR Info: " + inf + "\n" + "bar1: " + bar1 + "\n" + "bar2: " + bar2 + "\n" + "bar3: " + bar3);
        }
        private void dp_updateProgressBarsMaximum(object sender, string inf, int bar1, int bar2, int bar3) // negative values do not change indications
        {
            if (bar1 >= 0) pb_ReadData.Maximum = bar1;
            if (bar2 >= 0) pb_ConvertData.Maximum = bar2;
            if (bar3 >= 0) pb_SendData.Maximum = bar3;
            //MessageBox.Show("Form: " + Name + "\n" + "Sender: " + sender.ToString()+ "\n" + "MAX Info: " + inf + "\n" + "bar1: " + bar1 + "\n" + "bar2: " + bar2 + "\n" + "bar3: " + bar3);
        }
        private void cb_Detailed_APPLog_CheckedChanged(object sender, EventArgs e)
        {
            global.Data.F_LOGAPP_EXT = cb_LogEXT.Checked; // possibly obsolette
            initiateLogging();
        }
        private void cb_Update_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_Update.Checked) MessageBox.Show("This feature was programmed \nallthough it was NOT tested under all conditions.\nYou may use it 'AS IS'.\nAny feedback is appreciated", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
        private void cb_Delete_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_Delete.Checked) MessageBox.Show("This feature was programmed \nallthough it was NOT tested under all conditions.\nYou may use it 'AS IS'.\nAny feedback is appreciated", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
        private void cb_Skip_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_Skip.Checked) MessageBox.Show("This feature was programmed \nallthough it was NOT tested under all conditions.\nYou may use it 'AS IS'.\nAny feedback is appreciated", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }        
        private void TB_PREF_TextChanged(object sender, EventArgs e)
        {

        }
        private void cb_HasHeaders_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_HasHeaders.Checked == false)
            {
                MessageBox.Show("All files must have headers for columns.\r\nOtherwise operations may fail!\r\nClick 'OK' and follow the instructions below:\r\n" +
               "1. Read your file without headers and save it. Numeric headers will be added automatically", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
        }
        private void btn_FullAuto_Click(object sender, EventArgs e)
        {
            ReadInputFile(this.TB_INPUTFILE.Text, cb_HasHeaders.Checked);
            CONVERT_Start();
            APILogin(this.TB_API_LOGIN_SERVER.Text, this.API_LOGIN_NAME.Text, this.API_LOGIN_PWD.Text);
            cb_TestRun.Checked = false;
            cb_AutoRun.Checked = true;
            cb_Add.Checked = true;
            //string option = cbb_Job.SelectedItem.ToString();
            SendData(TB_OUTPUT_TABLE.Text, "ACE API Cmd");
        }
        private void cb_TestRun_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void lblLoginStatus2_Click(object sender, EventArgs e)
        {
            APILogin(TB_API_LOGIN_SERVER.Text, API_LOGIN_NAME.Text, API_LOGIN_PWD.Text);
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://www.youtube.com/channel/UCXxdkFr7eQIIhasfaPKqR1A");
        }
        private void TB_RAMLOGLENGTH_TextChanged(object sender, EventArgs e)
        {
            initiateLogging();
        }
        private void cb_LogSRV_CheckedChanged(object sender, EventArgs e)
        {
            initiateLogging();
        }
        private void cb_LogRAM_CheckedChanged(object sender, EventArgs e)
        {
            initiateLogging();
            if (!cb_LogRAM.Checked) dp.dumpRAMLog(9);                                                 // dump RAM Log to HDD if RAM logging was enabled
        }
        private void btn_RAMLog_Test_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= 2.5 * dp.RAM_LOGLength; i++)
            {
                msg = i.ToString() + " Test0; Test1; Test2; Test3; Test4; Test5";
                dp.logger(msg, "11000010");
            }
            dp.dumpRAMLog(9);                                                 // dump RAM Log to HDD if RAM logging was enabled
        }
        private void cb_LogOverwrite_CheckedChanged(object sender, EventArgs e)
        {
            dp.LOG_Overwrite = cb_LogOverwrite.Checked;
        }

#endregion-------------

        

    }
}


//LINKS TO OPEN RESOURCES
//http://www.cyberforum.ru/csharp-beginners/thread936775.html - как привязать DataTable к DataGridView и обновлять их в паре;
//https://msdn.microsoft.com/ru-ru/library/aka44szs(v=vs.110).aspx - подстрока;
//http://www.cyberforum.ru/windows-forms/thread737841.html - делегат
//http://blog.vkuznetsov.ru/posts/2011/08/27/csharp-net-eshhe-pyat-malenkih-chudes-kotorye-delayut-kod-luchshe-chast-2-iz-3#.VuRfMmxf13A - 5 нужных вещей
// String.IsNullOrEmpty("some text")
// http://metanit.com/sharp/adonet/2.9.php
// http://www.cyberforum.ru/ado-net/thread1447557.html
// https://otvet.mail.ru/question/69860826 - статика
//https://msdn.microsoft.com/ru-ru/library/8kb3ddd4(v=vs.110).aspx - дата время
// aceAuthorization.Dispose()
// Console.WriteLine("{0} was copied to {1}.", Data1, Data2);
// https://msdn.microsoft.com/ru-ru/library/dwhawy9k(v=vs.110).aspx стандартные форматы
// http://www.yandex.ru 