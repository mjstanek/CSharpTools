using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Identity.Client;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        try
        {
           // Console.WriteLine("Starting Dataverse Attribute Mapping Remover...");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //Console.WriteLine("Please select an environment to connect to.");
            using var selectForm = new DataverseMappingRemover.UI.Forms.SelectEnvironmentForm();
            var dialogResult = selectForm.ShowDialog();

            if (dialogResult != DialogResult.OK || selectForm.SelectedService is null)
            {
                MessageBox.Show("No environment was selected. Exiting application.", "Exit",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            var inputForm = new DataverseMappingRemover.UI.Forms.MappingLookupForm(selectForm.SelectedService)
            {
                TopMost = true,
                StartPosition = FormStartPosition.CenterScreen
            };
            var inputResult = inputForm.ShowDialog();
            if (inputResult == DialogResult.OK)
            {
                var entity = inputForm.SourceEntityLogicalName;
                var attribute = inputForm.SourceAttributeLogicalName;
                //Console.WriteLine($"Searching for attribute mappings for Entity: \"{entity}\", Attribute: \"{attribute}\"");
            }
            else
            {
               // Console.WriteLine("Exiting application.");
                return;
            }
        }
        finally
        {
           // Console.WriteLine("Closing application");
            Environment.Exit(0);
        }
    }
}