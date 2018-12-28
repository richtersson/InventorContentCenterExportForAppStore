using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Inventor;

namespace InventorContentCenterExport
{
    public class ContentCenterConnection
    {
        private bool InventorStarted = false;
        private Inventor.Application inventorApp = null;
        private List<Family> _families;
        private ObservableCollection<TreeItemViewModel> _menuitems;
        private TreeItemViewModel _root;
        private string _contentcenterpath;
        private string _generatedFileName;
        private int _filesCount;
        private string InventorProgID = "Inventor.Application";

        public Inventor.Application InventorApp { get => inventorApp; set => inventorApp = value; }
        public List<Family> Families { get => _families; set => _families = value; }
        public ObservableCollection<TreeItemViewModel> Menuitems { get => _menuitems; set => _menuitems = value; }
        public TreeItemViewModel Root { get => _root; set => _root = value; }
        public string Contentcenterpath { get => _contentcenterpath; set => _contentcenterpath = value; }
        public string GeneratedFileName { get => _generatedFileName; set => _generatedFileName = value; }
        public int FilesCount { get => _filesCount; set => _filesCount = value; }

        public ContentCenterConnection()
        {
            
            _menuitems = new ObservableCollection<TreeItemViewModel>();
            _root = new TreeItemViewModel(null, false) { DisplayName = "Content Center" };
            _root.Remarks = "Exporting..." + System.Environment.NewLine;
            _generatedFileName = "";
            _filesCount = 0;
            // Try to get an active instance of Inventor
            try
            {
                InventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject(InventorProgID) as Inventor.Application;
                
            }
            catch
            {
                //throw new Exception("Spustite Inventor a nastavte projekt pre export.");
            }
            if (inventorApp != null)
            {
                inventorApp = null;
                throw new Exception("Close Inventor and start program again.");
            }

            //If not active, create a new Inventor session
            if (null == inventorApp)
            {
                inventorApp = RestartInventor(inventorApp);
            }
        }
        public void ReadContentCenter(ContentTreeViewNode Node, TreeItemViewModel parentMenuItem)
        {
            foreach (ContentTreeViewNode oNode in Node.ChildNodes)
            {
                TreeItemViewModel oItem1 = new TreeItemViewModel(parentMenuItem,false);
                oItem1.DisplayName = oNode.DisplayName;
                parentMenuItem.Children.Add(oItem1);
                if (oNode.ChildNodes.Count > 0)
                {
                    ReadContentCenter(oNode, oItem1);
                }
            }

        }
        public void GenerateFiles(TreeItemViewModel inputnode)
        {
            //read tree
            TreeItemViewModel changingNode = inputnode;
            List<string> NodeNames = new List<string>();
            NodeNames.Add(changingNode.DisplayName);
            while (changingNode.Parent != null)
            {
                NodeNames.Add(changingNode.Parent.DisplayName);
                changingNode = changingNode.Parent;
            }

            ContentCenter oContentCenter = inventorApp.ContentCenter;
            ContentTreeViewNode ChosenNode;
            ChosenNode = oContentCenter.TreeViewTopNode;


            for (int i = 2; i <= NodeNames.Count; i++)
            {
                ChosenNode = ChosenNode.ChildNodes[NodeNames[NodeNames.Count - i]];

            }

            if (ChosenNode.Families.Count > 0)
            {
                GenerateFiles(ChosenNode);
            }
            if (ChosenNode.ChildNodes.Count > 0)
            {
                GetFamiliesAndGenerateFiles(ChosenNode);
            }
        }

        public void GetFamiliesAndGenerateFiles(ContentTreeViewNode oNode)
        {
           
            
                foreach (ContentTreeViewNode node in oNode.ChildNodes)
                {
                    if (node.Families.Count>0)
                {
                    GenerateFiles(node);
                }
                    if (node.ChildNodes.Count>0)
                {
                    GetFamiliesAndGenerateFiles(node);
                }
                }
            
        }
        public void GenerateFiles(ContentTreeViewNode oNode)
        {

            
            MemberManagerErrorsEnum error;
            string strContentPartFileName;
            string strErrorMessage;
            foreach (ContentFamily oFamily in oNode.Families)
            {
                foreach (ContentTableRow oRow in oFamily.TableRows)
                {
                    strContentPartFileName = oFamily.CreateMember(oRow, out error, out strErrorMessage);
                    _filesCount++;
                    using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(System.IO.Path.GetTempPath() +@"\InventorContentCenterExportLog.txt", true))
                    {
                        file.WriteLine(strContentPartFileName + "|||" + error + "|||" + strErrorMessage);
                    }
                    _root.Remarks = _root.Remarks + strContentPartFileName + System.Environment.NewLine;
                    if (_filesCount >= 50)
                    {
                        inventorApp.Documents.CloseAll();
                        //inventorApp = RestartInventor(inventorApp);
                        _filesCount = 0;
                    }
                }

            }
            

        }

        private Inventor.Application RestartInventor(Inventor.Application application)
        {
            if (application != null)
            {
                try
                {
                    application.Quit();
                }
                catch
                {

                }
            }

            Inventor.Application _InventorApp;
            Type inventorAppType = Type.GetTypeFromProgID(InventorProgID); // giving exception in Admin mode

            if (null != inventorAppType)
            {

               _InventorApp = (Inventor.Application)System.Activator.CreateInstance(inventorAppType) as Inventor.Application;

                if (null != _InventorApp)
                {
                    // set the flag to indicate we started the Inventor app so we shall quit it.
                    InventorStarted = true;
                    System.Threading.Thread.Sleep(3000);
                }
            }
            else
            {

                throw new Exception("Inventor not installed in the computer.");
            }
            return _InventorApp;

        }

        public void GenerateFilesOnSecondThread(TreeItemViewModel[] array)
        {
            //Thread newWindowThread = new Thread(new ThreadStart(() =>
            //{
            //AutoResetEvent are = (AutoResetEvent)threadclass.WaitHandles[0];
            foreach (TreeItemViewModel node in array)
            {


                GenerateFiles(node);

            }
            int numLines = _root.Remarks.Split('\n').Length;
            _root.Remarks = _root.Remarks + "Export finished." + (numLines-2) + " files were exported.";
            MessageBox.Show("Selected files were succesfully generated.");

           
            //    System.Windows.Threading.Dispatcher.Run();
            //}));
            //// set the apartment state  
            //newWindowThread.SetApartmentState(ApartmentState.STA);

            //// make the thread a background thread  
            //newWindowThread.IsBackground = true;

            //// start the thread  
            //newWindowThread.Start();
            //newWindowThread.Join();
        }


    }
    public class FamilyMember
    {
        private string _fileName;
        private string _partNumber;
        private string _description;

        public string FileName { get => _fileName; set => _fileName = value; }
        public string PartNumber { get => _partNumber; set => _partNumber = value; }
        public string Description { get => _description; set => _description = value; }
    }
    public class Family
    {
        private string _name;
        private List<FamilyMember> _members;

        public Family()
        {
            _members = new List<FamilyMember>();
        }

        public string Name { get => _name; set => _name = value; }
        public List<FamilyMember> Members { get => _members; set => _members = value; }
    }

    public class MenuItem 
    {
        public MenuItem()
        {
            this.Items = new ObservableCollection<MenuItem>();
        }

        public string Title { get; set; }

        public ObservableCollection<MenuItem> Items { get; set; }
    }
}
