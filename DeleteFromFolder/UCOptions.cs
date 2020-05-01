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
using Microsoft.Office.Interop.Outlook;

namespace DeleteFromFolder
{
    [ComVisible(true)]
    public partial class UCOptions : UserControl, PropertyPage
    {
        [DispId(-518)]
        public string Caption => "Delete From Folder";
                
        private PropertyPageSite _PropertyPageSite { get; set; }
        private Stores _Stores { get; set; }
        private bool _bDirty { get; set; }
        
        public UCOptions(Stores stores)
        {
            InitializeComponent();

            this._Stores = stores;
            this._PropertyPageSite = null;
            this._bDirty = false;
        }

        #region PropertyPage Methods

        public void Apply()
        {
            if (_bDirty)
            {
                SaveOptions();
                OnDirty(false);
            }
        }
        public bool Dirty => _bDirty;
        public void GetPageInfo(ref string HelpFile, ref int HelpContext)
        {
            MessageBox.Show("No Help available", "Warning", MessageBoxButtons.OK);
        }

        #endregion

        private PropertyPageSite GetPropertyPageSite()
        {
            Type type = typeof(object);
            string assembly = type.Assembly.CodeBase.Replace("mscorlib.dll", "System.Windows.Forms.dll");
            assembly = assembly.Replace("file:///", "");

            string assemblyName = System.Reflection.AssemblyName.GetAssemblyName(assembly).FullName;
            Type unsafeNativeMethods = Type.GetType(System.Reflection.Assembly.CreateQualifiedName(assemblyName, "System.Windows.Forms.UnsafeNativeMethods"));

            Type oleObj = unsafeNativeMethods.GetNestedType("IOleObject");
            System.Reflection.MethodInfo methodInfo = oleObj.GetMethod("GetClientSite");
            object propertyPageSite = methodInfo.Invoke(this, null);

            return (PropertyPageSite)propertyPageSite;
        }
        private void OnDirty(bool isDirty)
        {
            _bDirty = isDirty;

            if(_PropertyPageSite != null)
                _PropertyPageSite.OnStatusChange();
        }
        private void LoadOptions()
        {
            //fill checkedlistbox with folders
            foreach (Store s in this._Stores)
            {
                MAPIFolder mapi = s.GetRootFolder();

                foreach (Folder f in mapi.Folders)
                    clbFolders.Items.Add(f.Name, Properties.Settings.Default.CheckedItems.Contains(f.Name));
            }
        }
        private void SaveOptions()
        {
            try
            {
                foreach (string checkedName in clbFolders.CheckedItems)
                    Properties.Settings.Default.CheckedItems.Add(checkedName);

                Properties.Settings.Default.Save();
            }
            catch(System.Exception)
            {
                throw;
            }
        }
        private void UCOptions_Load(object sender, EventArgs ar)
        {
            LoadOptions();

            this._PropertyPageSite = GetPropertyPageSite();
        }
        private void clbFolders_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            OnDirty(true);
        }
    }
}
