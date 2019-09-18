using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DeleteFromFolder
{
    [ComVisible(true)]
    public partial class UCOptions : UserControl, Outlook.PropertyPage
    {
        #region Public Members

        [DispId(-518)]
        public string Caption
        {
            get
            {
                return "Delete From Folder";
            }
        }

        #endregion

        #region Private Members

        Outlook.PropertyPageSite _PropertyPageSite = null;
        bool _bDirty = false;
        Outlook.Stores _Stores;
        bool _bFlag = false;

        #endregion

        public UCOptions(Outlook.Stores stores)
        {
            InitializeComponent();

            this._Stores = stores;
                        
            this.Load += new EventHandler(UCOptions_Load);            
        }

        #region Public Methods

        public void Apply()
        {
            if (_bDirty)
            {
                SaveOptions();
                OnDirty(false);
            }
        }
        public bool Dirty
        {
            get { return _bDirty; }
        }
        public void GetPageInfo(ref string HelpFile, ref int HelpContext)
        {
            MessageBox.Show("No Help available", "Warning", MessageBoxButtons.OK);
        }

        #endregion

        #region Private Methods

        void UCOptions_Load(object sender, EventArgs ar)
        {
            LoadOptions();

            _PropertyPageSite = GetPropertyPageSite();
        }
        Outlook.PropertyPageSite GetPropertyPageSite()
        {
            Type type = typeof(object);
            string assembly = type.Assembly.CodeBase.Replace("mscorlib.dll", "System.Windows.Forms.dll");
            assembly = assembly.Replace("file:///", "");

            string assemblyName = System.Reflection.AssemblyName.GetAssemblyName(assembly).FullName;
            Type unsafeNativeMethods = Type.GetType(System.Reflection.Assembly.CreateQualifiedName(assemblyName, "System.Windows.Forms.UnsafeNativeMethods"));

            Type oleObj = unsafeNativeMethods.GetNestedType("IOleObject");
            System.Reflection.MethodInfo methodInfo = oleObj.GetMethod("GetClientSite");
            object propertyPageSite = methodInfo.Invoke(this, null);

            return (Outlook.PropertyPageSite)propertyPageSite;
        }
        void OnDirty(bool isDirty)
        {
            _bDirty = isDirty;

            _PropertyPageSite.OnStatusChange();
        }
        void LoadOptions()
        {
            //fill checkedlistbox with folders
            foreach (Outlook.Store s in this._Stores)
            {
                Outlook.MAPIFolder mapi = s.GetRootFolder();

                foreach (Outlook.Folder f in mapi.Folders)
                {
                    clbFolders.Items.Add(f.Name);
                }
            }

            //set the ones that are checked
            foreach (string s in DeleteFromFolder.Properties.Settings.Default.CheckedItems.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int index = clbFolders.FindString(s);

                if (index >= 0)
                    clbFolders.SetItemChecked(index, true);
            }

            _bFlag = true; //this is set to false initially so that ItemCheck event doesnt run while we are populating list with saved data
        }
        void SaveOptions()
        {
            try
            {
                string idx = string.Empty;
                foreach (string s in (from object l in clbFolders.CheckedItems select l.ToString()).ToArray())
                {
                    idx += (string.IsNullOrEmpty(idx) ? string.Empty : ",") + s;
                }

                DeleteFromFolder.Properties.Settings.Default.CheckedItems = idx;
                DeleteFromFolder.Properties.Settings.Default.Save();
            }
            catch(Exception)
            {
                throw;
            }
        }

        #endregion

        private void clbFolders_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if(_bFlag)
                OnDirty(true);
        }
    }
}
