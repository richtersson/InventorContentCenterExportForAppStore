using System;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace InventorContentCenterExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private ContentCenterConnection _oConnection;
        private TreeItemViewModel _root;



        public TreeItemViewModel Root
        {
            get
            {
                return _root;
            }
            set
            {
                if (_root == value)
                {
                    return;
                }
                _root = value;
                //RaisePropertyChanged(PartsListCollection);
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            this.Title = "Generovanie súčiastok Obsahového centra Inventoru " + AddInGlobal.AppVersion;


        }
        void RaisePropertyChanged(string prop)
        {
            if (PropertyChanged != null) { PropertyChanged(this, new PropertyChangedEventArgs(prop)); }
        }
        public event PropertyChangedEventHandler PropertyChanged;


        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        private void btnLoadContentCenter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _oConnection = new ContentCenterConnection();
                
                _oConnection.ReadContentCenter(_oConnection.InventorApp.ContentCenter.TreeViewTopNode, _oConnection.Root);
                //TreeItemViewModel MainNode = new TreeItemViewModel();
                //MainNode.Children = oConnection.Menuitems;
                //MainNode.Title = "Obsahove centrum";
                //trvMenu.Items.Add(MainNode);
                Root = _oConnection.Root;
                DataContext = Root;
                btnGenerateFiles.IsEnabled = true;
                txtDirectory.Text = _oConnection.InventorApp.DesignProjectManager.ActiveDesignProject.ContentCenterPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnGetDirectory_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string path = dialog.SelectedPath;

                txtDirectory.Text = path;
            }
        }

        private void btnGenerateFiles_Click(object sender, RoutedEventArgs e)
        {
            txtFiles.Visibility = Visibility.Visible;
            brFiles.Visibility = Visibility.Visible;
            this.InvalidateVisual();
            this.Refresh();
            try
            {

                TreeItemViewModel[] array = TheTreeView.SelectedItems.OfType<TreeItemViewModel>().ToArray();

                var tasks = new Task[1];

                tasks[0] = Task.Factory.StartNew(() => _oConnection.GenerateFilesOnSecondThread(array));                
                //Task.WaitAll(tasks);
                
                


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                //_oConnection = null;
                //btnGenerateFiles.IsEnabled = false;

            }

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            _oConnection?.InventorApp?.Quit();
            //if (_oConnection.InventorApp != null)
            //    _oConnection.InventorApp.Quit();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            About oAbout = new About();
            oAbout.ShowDialog();
        }
    }
    public static class ExtensionMethods
    {
        private static Action EmptyDelegate = delegate () { };


        public static void Refresh(this UIElement uiElement)

        {
            uiElement.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);
        }
    }
    public static class AutoScrollBehavior
    {
        public static readonly DependencyProperty AutoScrollProperty =
            DependencyProperty.RegisterAttached("AutoScroll", typeof(bool), typeof(AutoScrollBehavior), new PropertyMetadata(false, AutoScrollPropertyChanged));


        public static void AutoScrollPropertyChanged(DependencyObject obj, DependencyPropertyChangedEventArgs args)
        {
            var scrollViewer = obj as ScrollViewer;
            if (scrollViewer != null && (bool)args.NewValue)
            {
                scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
                scrollViewer.ScrollToEnd();
            }
            else
            {
                scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
            }
        }

        private static void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // Only scroll to bottom when the extent changed. Otherwise you can't scroll up
            if (e.ExtentHeightChange != 0)
            {
                var scrollViewer = sender as ScrollViewer;
                scrollViewer?.ScrollToBottom();
            }
        }

        public static bool GetAutoScroll(DependencyObject obj)
        {
            return (bool)obj.GetValue(AutoScrollProperty);
        }

        public static void SetAutoScroll(DependencyObject obj, bool value)
        {
            obj.SetValue(AutoScrollProperty, value);
        }
    }
    public  class ThreadClass
    {
        // Define an array with two AutoResetEvent WaitHandles.
        private  WaitHandle[] waitHandles;
        private TreeItemViewModel[] array;
        public ThreadClass(TreeItemViewModel[] _array)
        {
            WaitHandles = new WaitHandle[]
                {
                new AutoResetEvent(false)

                };
            array = _array;
        }

        public WaitHandle[] WaitHandles { get => waitHandles; set => waitHandles = value; }
        public TreeItemViewModel[] Array { get => array; set => array = value; }
    }

}
