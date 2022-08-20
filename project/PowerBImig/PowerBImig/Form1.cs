using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;


namespace PowerBImig
{


    public partial class Form1 : Form
    {
    
        /// ----------------------------------------------------------------------------------------------------------------------------------------------------------------- ///
        ///
        /// Global Constants and Variables
        ///    

        // Constants
        const string HTTPHEADUSERAGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36";
        const string PBI_API_URLBASE = "https://api.powerbi.com/v1.0/myorg/";

        const string AuthorityURL = "https://login.windows.net/common/oauth2/authorize";  // use with 3.19.8

        const string ResourceURL = "https://analysis.windows.net/powerbi/api";
        const string RedirectURL = "https://dev.powerbi.com/Apps/SignInRedirect";
        // Native Azure AD App ClientID  --  Put your Client ID here
        public static string ApplicationID;

        // Variables 
        HttpClient client = null;


        /// ----------------------------------------------------------------------------------------------------------------------------------------------------------------- ///
        ///
        /// Default Class constructor
        ///    
      

        /// ----------------------------------------------------------------------------------------------------------------------------------------------------------------- ///
        ///
        /// Execute Method
        ///    
        public void Execute(string authType)
        {
            string authToken = "";


            // Get an authentication token
            authToken = GetAuthTokenUser(authType);  // Uses native AD auth

            // Initialize the client with the token
            if (!String.IsNullOrEmpty(authToken))
            {
                InitHttpClient(authToken);

                GetWorkspaces();

         
            }
        }

        public string GetAuthTokenUser(string authType)
        {
            Task<AuthenticationResult> authResult = null;
            string authToken = "";
            authResult = GetAuthUserLoginInteractive();

            return authToken;

        }

            public void GetWorkspaces()
        {

        }


            public void InitHttpClient(string authToken)
        {
         

            // Create the web client connection
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd(HTTPHEADUSERAGENT);
            client.DefaultRequestHeaders.Add("Authorization", authToken);
        }
        /// ----------------------------------------------------------------------------------------------------------------------------------------------------------------- ///
        ///
        /// GetAuthUserLogin Method (Interactive)
        /// 
        public async Task<AuthenticationResult> GetAuthUserLoginInteractive()
        {
            AuthenticationResult authResult = null;

            PlatformParameters parameters = new PlatformParameters(PromptBehavior.Always);

            try
            {
                // Query Azure AD for an interactive login prompt and subsequent Power BI auth token
                AuthenticationContext authContext = new AuthenticationContext(AuthorityURL);
                authResult = await authContext.AcquireTokenAsync(ResourceURL, ApplicationID, new Uri(RedirectURL), parameters).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
               

                lbloutput.Text = "Error acquiring token with interactive credential" + ex.Message;

                authResult = null;
            }

            return authResult;
        }


       


            public Form1()
            {
                InitializeComponent();
            }

            public void button1_Click(object sender, EventArgs e)
        {
               
                
                lbloutput.Text = ApplicationID;
               

        }

            public void textBox1_TextChanged(object sender, EventArgs e)
            {

            InitializeComponent();

            ApplicationID = textBox1.Text;

        }

            private void Form1_Load(object sender, EventArgs e)
            {


            }




            private void label2_Click(object sender, EventArgs e)
            {

            



        }

        private void button2_Click(object sender, EventArgs e)
        {
            //     PowerShell ps = PowerShell.Create();
            //  ps.AddCommand("Connect-PowerBIServiceAccount").Invoke();
            //  ps.AddCommand("Get-PowerBIAccessToken");


            //var results = ps.Invoke();

            PowerShell ps = PowerShell.Create();
           // ps.AddCommand("Connect-PowerBIServiceAccount").Invoke();

          //  lbloutput.Text = plogin("Connect-PowerBIServiceAccount");

           
            textBox2.Text = plogin("Get-Process");


        }

        private string plogin(string cmd)
        {

                 
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript(cmd);
            pipeline.Commands.Add("out-string");
            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();

            StringBuilder stringBuilder = new StringBuilder();
            foreach(PSObject pSObject in results)
                stringBuilder.AppendLine(pSObject.ToString());
            return results.ToString();

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           
        }
    }
    }
