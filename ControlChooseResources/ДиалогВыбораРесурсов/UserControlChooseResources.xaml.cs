//using ControlChooseResources;
using System;
using System.Collections.Generic;
//using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
//using System.Windows.Input;

namespace ControlChooseResources
{
    /// <summary>
    /// Interaction logic for UserControlChooseResources.xaml
    /// </summary>
    public partial class UserControlChooseResources : UserControl
    {
        public UserControlChooseResources()
        {
            InitializeComponent();
            Root = new Resource(null, null, DateTime.Now, DateTime.MinValue);
            treeView.ItemsSource = new List<TreeViewNode>() { new TreeViewNode(Root) };
            foreach (var item in treeView.Items)
            {
                (item as TreeViewNode).UncheckAllNotActualChildren(((DateTime)dateActual.SelectedDate).Date);
            }
        }



        public List<Resource> CheckedResources
        { get {
                List<Resource> result = new List<Resource>();

                foreach (var item in treeView.Items)
                    result.AddRange((item as TreeViewNode).GetAllCheckedResources());

                /*StringBuilder str = new StringBuilder();
                foreach (var item in result)
                {
                    str.AppendLine(item.Parent.Name + " - " + item.Name);
                }
                MessageBox.Show(str.ToString());
                */return result;
            } }

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

        private void onlyActual_Checked(object sender, RoutedEventArgs e)
        {
            if (treeView == null) return;

            foreach (var item in treeView.Items)
            {
                (item as TreeViewNode).OnlyActual = true;
                (item as TreeViewNode).UncheckAllNotActualChildren(((DateTime)dateActual.SelectedDate).Date);
            }

        }

        private void onlyActual_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void checkBox_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {
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
