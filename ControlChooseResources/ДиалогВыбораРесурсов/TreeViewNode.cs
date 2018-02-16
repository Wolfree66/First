using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace ControlChooseResources
{
    class TreeViewNode : FrameworkElement
    {

        TreeViewNode()
        {
        }
        #region Статическая часть


        static TreeViewNode()
        {
            _Data = DateTime.Today;
            _OnlyActual = false;
            Inctance = new TreeViewNode();
        }


        public static TreeViewNode Inctance { get; private set; }


        #endregion

        // свойство зависимостей
        // public static readonly DependencyProperty OnlyActualProperty =
        //   DependencyProperty.Register("OnlyActual", typeof(bool), typeof(TreeViewNode), new FrameworkPropertyMetadata(null));

        static bool _OnlyActual;
        public bool OnlyActual
        {
            get { return _OnlyActual; }
            set
            {
                if (value == _OnlyActual) return;
                _OnlyActual = value;
                NotifyPropertyChanged("OnlyActual");
            }
        }

        public TreeViewNode(Resource res)
        {
            this.Item = res;
            //Data = DateTime.Today;
            IsActual = IsActualOnDate(Data);
            //this.IsChecked = IsActual;

        }

        static DateTime _Data;
        public DateTime Data
        {
            get { return _Data; }
            set
            {
                if (value == _Data) return;
                _Data = value;
            }
        }

        Resource _Item;
        public Resource Item
        {
            get { return _Item; }
            set
            {
                if (_Item == value) return;
                _Item = value;
                _NodeName = Item.Name;
                List<TreeViewNode> children = new List<TreeViewNode>();
                foreach (var item in Item.Children)
                    children.Add(new TreeViewNode(item));
                Children = children;

            }
        }

        // свойство зависимостей
        public static readonly DependencyProperty IsActualProperty =
          DependencyProperty.Register("IsActual", typeof(bool), typeof(TreeViewNode), new FrameworkPropertyMetadata(null));

        public bool IsActual
        {
            get { return (bool)GetValue(IsActualProperty); }
            set
            {
                SetValue(IsActualProperty, value);
                NotifyPropertyChanged("IsActual");
                //this.SetIsChecked(value, true, true);
            }
        }

        public bool IsActualOnDate(DateTime date)
        {
            bool result = false;
            if (Item.EndUsing == null || Item.EndUsing == DateTime.MinValue) result = true;
            if (Item.EndUsing >= date.Date) result = true;
            IsActual = result;
            return result;
        }

        // свойство зависимостей
        public static readonly DependencyProperty IsCheckedProperty =
            DependencyProperty.Register("IsChecked", typeof(bool), typeof(TreeViewNode), new FrameworkPropertyMetadata(false, new PropertyChangedCallback(OnIsCheckedValuePropertyChanged), new CoerceValueCallback(IsCheckedCoerceValue)));

        private static void OnIsCheckedValuePropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            TreeViewNode instance = (TreeViewNode)d;
            bool newPropertyValue = (bool)e.NewValue;
            if (instance.OnlyActual && instance.IsActual)
                instance.IsChecked = newPropertyValue;
            else if (!instance.OnlyActual) instance.IsChecked = newPropertyValue;
            else return;
            instance.NotifyPropertyChanged();
        }

        private static object IsCheckedCoerceValue(DependencyObject obj, object baseValue)
        {
            // cast the object as a login control
            bool result = true;
            TreeViewNode tvn = (TreeViewNode)obj;
            Double hours = 0;
            bool newValue = (bool)baseValue;
            /*if (tvn.IsActual) result = newValue;
            else */
            if (tvn.OnlyActual && !tvn.IsActual) result = false;
            else result = newValue;
            return result;
        }

        /// <summary>
        /// Gets/sets the state of the associated UI toggle (ex. CheckBox).
        /// The return value is calculated based on the check state of all
        /// child FooViewModels.  Setting this property to true or false
        /// will set all children to the same check state, and setting it 
        /// to any value will cause the parent to verify its check state.
        /// </summary>
        public bool IsChecked
        {
            get { return (bool)GetValue(IsCheckedProperty); }
            set
            {
                SetValue(IsCheckedProperty, value);
                CheckAllChildren((bool)value);
                NotifyPropertyChanged();
                //this.SetIsChecked(value, true, true);
            }
        }

        public void ActualiseAllChildren()
        {
            this.IsActualOnDate(Data);
            foreach (var item in Children)
            {
                //item.IsActualOnDate(Data);
                item.ActualiseAllChildren();
            }
        }

        public void CheckAllChildren(bool check)
        {
            foreach (var item in Children)
            {
              //  if (OnlyActual && item.IsActual)
                    item.IsChecked = check;
               // else if(OnlyActual && !item.IsActual) ;
                item.CheckAllChildren(check);
            }
        }

        public void UncheckAllNotActualChildren(DateTime date)
        {
            foreach (var item in Children)
            {
                if (!item.IsActualOnDate(date))
                    (item as TreeViewNode).IsChecked = false;
                item.UncheckAllNotActualChildren(date);
            }
        }

        public List<Resource> GetAllCheckedResources()
        {
            List<Resource> result = new List<Resource>();
            if (this.IsChecked)
            {
                if (OnlyActual && this.IsActual) result.Add(this.Item);
                else if (!OnlyActual) result.Add(this.Item);
            }
            foreach (var item in Children)
            {
                result.AddRange(item.GetAllCheckedResources());
            }
            return result;
        }

        CollectionViewSource _Items;
        public CollectionViewSource Items
        {
            get
            {
                if (_Items != null) return _Items;
                CollectionViewSource cvs = new CollectionViewSource();
                cvs.Source = Children;
                _Items = cvs;
                return cvs;
            }
            set
            {
                if (value == null || value == _Items) return;
                _Items = value;
            }
        }

        public List<TreeViewNode> Children
        { get; set; }


        string _NodeName;
        public string NodeName
        { get { return _NodeName; } }


        TreeViewNode _parent;


        #region 
        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property.
        // The CallerMemberName attribute that is applied to the optional propertyName
        // parameter causes the property name of the caller to be substituted as an argument.
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion
    }
}

