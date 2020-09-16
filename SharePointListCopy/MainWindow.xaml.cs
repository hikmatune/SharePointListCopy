using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.SharePoint.Client;
using List = Microsoft.SharePoint.Client.List;
using OfficeDevPnP.Core;

namespace SharePointListCopy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            txtOutput.IsReadOnly = true;
            
        }

        private void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            ArrayList arrayList = new ArrayList();
            /// action on source url
            try
            {
                
                /// get authentication
                var authSourceManager = new OfficeDevPnP.Core.AuthenticationManager();

                var sourceContext = authSourceManager.GetWebLoginClientContext(txtSourceURL.Text);
                Web _web = sourceContext.Web;
                sourceContext.Load(_web);
                try
                {
                    sourceContext.ExecuteQuery();
                    txtOutput.Text += "Successfully connected to " + txtSourceURL.Text + "\n";

                }
                catch
                {
                    txtOutput.Text += "Failed to connect to " + txtSourceURL.Text + "\n";

                }

                List _list = sourceContext.Web.Lists.GetByTitle(txtListName.Text);
                FieldCollection _fields = _list.Fields;
                sourceContext.Load(_list.Fields);
                sourceContext.Load(_fields);
                try
                {
                    sourceContext.ExecuteQuery();
                }
                catch
                {
                    txtOutput.Text += "Faild to get the list " + txtListName.Text + "\n";
                    txtOutput.Text += "Please make sure the list exists \n";
                }
                txtOutput.Text += "Getting " + txtListName.Text + " custom fields.... \n";
                foreach (Field _field in _fields)
                {
                    if (_field.CanBeDeleted)
                    {

                        txtOutput.Text += "---" + _field.Title + "\n";
                        arrayList.Add(_field.SchemaXml);
                    }

                }
                txtOutput.Text += "There are " + arrayList.Count + " custom fields in " + txtListName.Text + "\n";

            }
            catch (Exception ex)
            {
                txtOutput.Text += "Exception happended " + ex.Message + "\n";
            }

            /// actions on destination url
            try
            {
                var authDetinationManager = new OfficeDevPnP.Core.AuthenticationManager();
                var destinationContext = authDetinationManager.GetWebLoginClientContext(txtDestURL.Text);
                Web destWeb = destinationContext.Web;
                ListCreationInformation creatListInfo = new ListCreationInformation();
                creatListInfo.Title = txtListName.Text;
                creatListInfo.TemplateType = (int)ListTemplateType.GenericList;
                List newList = destWeb.Lists.Add(creatListInfo);
                if(arrayList.Count > 0)
                {
                    for (int x= 0; x < arrayList.Count; x++)
                    {
                        string currentField = (string)arrayList[x];
                        newList.Fields.AddFieldAsXml(currentField, true, AddFieldOptions.AddToDefaultContentType);
                    }
                }
                try
                {
                    destinationContext.ExecuteQuery();
                    txtOutput.Text += "Successfully created the list with the fields at " + txtDestURL.Text + "\n";
                    txtOutput.Text += " You must manually fix the lookup fields \n";
                }catch(Exception ex)
                {
                    txtOutput.Text += "Failed to create the list \n";
                    txtOutput.Text += ex.Message + "\n";
                }
            }
            catch (Exception ex)
            {
                txtOutput.Text += "Exception happended " + ex.Message + "\n";
            }

        }

        private void btnClearLog_Click(object sender, RoutedEventArgs e)
        {
            txtOutput.Text = "";
        }
    }
}
