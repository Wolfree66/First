using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using TFlex.DOCs.Model;

namespace ControlChooseResources
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        public Window1()
        {
            
            InitializeComponent();
            Root = new Resource(null, null, DateTime.Now, DateTime.MinValue);
            treeView.ItemsSource = new List<TreeViewNode>() { new TreeViewNode(Root) };
            foreach (var item in treeView.Items)
            {
                (item as TreeViewNode).UncheckAllNotActualChildren(((DateTime)dateActual.SelectedDate).Date);
            }

        }

        public List<Resource> ChosenResources { get; set; }
        private Resource Root { get; set; }
        private void trw_Products_Expanded(object sender, RoutedEventArgs e)
        {
            return;
            TreeViewItem item = (TreeViewItem)e.OriginalSource;
            item.Items.Clear();

           /* try
            {*/
               // MessageBox.Show(((Resource)(item.Tag)).Name + "\n");
                List<Resource> children = ((Resource)(item.Tag)).Children;
                foreach (Resource res in children)
                {
                    
                    TreeViewItem newItem = new TreeViewItem();
                    newItem.Tag = res;
                    newItem.Header = res.ToString();
                    // MessageBox.Show(res.Children.Count.ToString());
                    if (res.Children != null && res.Children.Count > 0)
                    {
                        newItem.Items.Add("*");
                    }
                    item.Items.Add(newItem);
                }
            /*}
            catch
            { throw;}*/
        }

        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            TextBox tb = sender as TextBox;
            Resource fvm = tb.GetBindingExpression(TextBox.TextProperty).ResolvedSource as Resource;
            //MessageBox.Show(fvm.ToString() + fvm.EndUsing.ToString() + "\n" + fvm.IsActual.ToString());
            MessageBox.Show(tb.TemplatedParent.DependencyObjectType.Id.ToString());
        }

        private void onlyActual_Checked(object sender, RoutedEventArgs e)
        {
            if (treeView == null) return;

            foreach (var item in treeView.Items)
            {
                (item as TreeViewNode).OnlyActual = true;
                (item as TreeViewNode).UncheckAllNotActualChildren(((DateTime)dateActual.SelectedDate).Date);
            }
            //UpdateResourcesTree(new List<FooViewModel>() { new FooViewModel(Root) });
            //ICollectionView view = CollectionViewSource.GetDefaultView(new List<FooViewModel>() { new FooViewModel(Root) });
            //(treeView.ItemsSource as ICollectionView).Refresh();
        }

        private void onlyActual_Unchecked(object sender, RoutedEventArgs e)
        {
            //(item as TreeViewNode).OnlyActual = false;
            // UpdateResourcesTree(new List<FooViewModel>() { new FooViewModel(Root) });
            //treeView.UpdateLayout();
            //ICollectionView view = CollectionViewSource.GetDefaultView(new List<FooViewModel>() { new FooViewModel(Root) });
            //(treeView.ItemsSource as ICollectionView).Refresh();
        }

        void UpdateResourcesTree(List<TreeViewNode> listFVM)
        {
            foreach (var item in listFVM)
            {
                item.Items.Filter -= Items_Filter;
                item.Items.Filter += Items_Filter;
                if (item.Children != null && item.Children.Count > 0)
                    UpdateResourcesTree(item.Children);
            }
            return;
    }


        void Items_Filter(object sender, FilterEventArgs e)
        {
            if (onlyActual.IsChecked == true)
            {
                DateTime date = (DateTime)dateActual.SelectedDate;
                TreeViewNode filteredNode = e.Item as TreeViewNode;
                e.Accepted = filteredNode.IsActualOnDate(date) == true;
            }
            else
                e.Accepted = true;
        }

        private void checkBox_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            
            //FrameworkElement.
            //Resource fvm = cb.GetBindingExpression(CheckBox.IsCheckedProperty).ResolvedSource as Resource;
            //TreeViewItem treeItem = this.treeView.ItemContainerGenerator.ContainerFromItem(fvm) as TreeViewItem;
            //MessageBox.Show(this.treeView.ItemContainerGenerator.ContainerFromItem(fvm).GetType().ToString());
            //FooViewModel fvm = (cb.Parent as TreeViewItem) as FooViewModel;
            //CheckAllChildren(fvm);
        }

        void CheckAllChildren(TreeViewNode tvi, bool onlyActual)
        {
            foreach (var item in tvi.Children)
            {
                if (onlyActual)
                {
                    if (item.IsActual)
                        (item as TreeViewNode).IsChecked = true;
                }
                CheckAllChildren(item, onlyActual);
            }
        }

        void UnCheckAllChildren(TreeViewNode tvi)
        {
            foreach (var item in tvi.Children)
            {
                (item as TreeViewNode).IsChecked = false;
                UnCheckAllChildren(item);
            }
        }

        private void dateActual_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (treeView == null) return;
            foreach (var item in treeView.Items)
            {
                MessageBox.Show((item as TreeViewNode).NodeName);
                (item as TreeViewNode).Data = (DateTime)dateActual.SelectedDate;
                (item as TreeViewNode).ActualiseAllChildren();
            }
            treeView.UpdateLayout();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            List<Resource> result = new List<Resource>();
            
            foreach (var item in treeView.Items)
             result.AddRange((item as TreeViewNode).GetAllCheckedResources());

            StringBuilder str = new StringBuilder();
            foreach (var item in result)
            {
                str.AppendLine(item.Parent.Name + " - "+ item.Name);
            }
            ChosenResources = result;
            MessageBox.Show(str.ToString());
        }

        private IEnumerable<TreeViewNode> GetCheckedItems(TreeViewNode node)
        {
            var checkedItems = new List<TreeViewNode>();

            ProcessNode(node, checkedItems);

            return checkedItems;
        }

        private void ProcessNode(TreeViewNode node, List<TreeViewNode> checkedItems)
        {
            foreach (var child in node.Children)
            {
                if ((bool)child.IsChecked)
                    checkedItems.Add(child);

                ProcessNode(child, checkedItems);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void hideNotActual_Checked(object sender, RoutedEventArgs e)
        {
            onlyActual.IsChecked = (sender as CheckBox).IsChecked;
        }

        private void hideNotActual_Unchecked(object sender, RoutedEventArgs e)
        {
           onlyActual.IsChecked = (sender as CheckBox).IsChecked;
        }

        private void treeView_Loaded(object sender, RoutedEventArgs e)
        {
            ItemContainerGenerator gen = treeView.ItemContainerGenerator;
            var tvi = gen.ContainerFromItem(treeView.Items[0]) as TreeViewItem;
            tvi.IsExpanded = true;
        }
    }
}
