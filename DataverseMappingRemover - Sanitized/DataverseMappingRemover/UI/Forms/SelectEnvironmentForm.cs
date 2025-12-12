using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using static System.Net.Mime.MediaTypeNames;
using DataverseMappingRemover.UI.Components;

namespace DataverseMappingRemover.UI.Forms
{
    public partial class SelectEnvironmentForm : Form
    {
        private RoundedButton? ProductionButton;
        private RoundedButton? SandboxButton;
        /*
         * UPDATE THESE URLS
        private const string productionUrl = "https://YOUR-DYNAMICS-PRODUCTION.crm.dynamics.com";
        private const string sandboxUrl = "https://YOUR-DYNAMICS-SANDBOX.crm.dynamics.com";
        */
        public enum EnvironmentOption
        {
            Sandbox,
            Production,
            Cancel
        }

        public EnvironmentOption Choice { get; private set; } = EnvironmentOption.Cancel;

        public Microsoft.Xrm.Sdk.IOrganizationService? SelectedService { get; private set; }

        public SelectEnvironmentForm()
        {
            Text = "Select an Environment";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(360, 140);

            var label = new System.Windows.Forms.Label
            {
                Text = "Select the environment you want to connect to:",
                AutoSize = true,
                Location = new Point(20, 20),
                TabIndex = 0,
                TabStop = false
            };
            Controls.Add(label);

            ProductionButton = new RoundedButton
            {
                Text = "Production",
                Location = new Point(40, 70),
                Size = new Size(120, 32),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(255, 15, 15),
                TabIndex = 1,
                TabStop = true
            };
            ProductionButton.Click += (_, __) =>
            {
                //Console.WriteLine("Connecting to Production...");
                Choice = EnvironmentOption.Production;
                var productionLogin = 
                    "AuthType=OAuth;" +
                    $"Url={productionUrl};" +
                    "AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;" +
                    "RedirectUri=http://localhost;" +
                    "LoginPrompt=Auto";
                var serviceClient = new Microsoft.PowerPlatform.Dataverse.Client.ServiceClient(productionLogin);
                SelectedService = serviceClient;
                DialogResult = DialogResult.OK;
               // Console.WriteLine("Connected to Production.");
                Close();
            };
            Controls.Add(ProductionButton);

            SandboxButton = new RoundedButton
            {
                Text = "Sandbox",
                Location = new Point(200, 70),
                Size = new Size(120, 32),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(76, 114, 41),
                TabIndex = 2,
                TabStop = true
            };
            SandboxButton.Click += (_, __) =>
            {
                //Console.WriteLine("Connecting to Sandbox...");
                Choice = EnvironmentOption.Sandbox;
                var sandboxLogin =
                    "AuthType=OAuth;" +
                    $"Url={sandboxUrl};" +
                    "AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;" +
                    "RedirectUri=http://localhost;" +
                    "LoginPrompt=Auto";
                var serviceClient = new Microsoft.PowerPlatform.Dataverse.Client.ServiceClient(sandboxLogin);
                SelectedService = serviceClient;
                DialogResult = DialogResult.OK;
                //Console.WriteLine("Connected to Sandbox.");
                Close();
            };
            Controls.Add(SandboxButton);

            ActiveControl = ProductionButton;
            AcceptButton = ProductionButton;
        }

        protected override bool ShowFocusCues => true;

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // System.Diagnostics.Debug.WriteLine($"Form ProcessCmdKey: {keyData}");

            if (keyData == Keys.Right || keyData == Keys.Down)
            {
                SelectNextControl(ActiveControl, true, true, true, true);
                return true;
            }
            else if (keyData == Keys.Left || keyData == Keys.Up)
            {
                SelectNextControl(ActiveControl, false, true, true, true);
                return true;
            }
            else if (keyData == Keys.Escape)
            {
                Choice = EnvironmentOption.Cancel;
                DialogResult = DialogResult.Cancel;
                Close();
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            if (ProductionButton is not null)
            {
                BeginInvoke(new Action(() =>
                {
                    //Console.WriteLine("Setting focus to ProductionButton");
                    ProductionButton.TabStop = true;
                    ProductionButton.Select();
                    ProductionButton.Focus();
                }));
            }
        }

    }
}
