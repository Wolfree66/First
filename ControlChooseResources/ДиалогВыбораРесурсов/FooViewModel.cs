using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace ДиалогВыбораРесурсов
{
    class FooViewModel : IFooViewModel
    {

        public FooViewModel(Resource res)
        {
            this.Item = res;
            Data = DateTime.Today;
            this.IsChecked = IsActual;
            
        }

        public static DateTime Data { get; set; }

        Resource _Item;
        public Resource Item {
            get { return _Item; }
            set {
                if (_Item == value) return;
                _Item = value;
                _Name = Item.Name;
                List<FooViewModel> children = new List<FooViewModel>();
                foreach (var item in Item.Children)
                    children.Add(new FooViewModel (item));
                Children = children;
                
            }
        }

        public bool IsActual
        {
            get
            {
                if (Item.EndUsing == null || Item.EndUsing == DateTime.MinValue) return true;
                if (Item.EndUsing >= Data) return true;
                return false;
            }
        }

        public bool IsActualOnDate(DateTime date)
        {
            if (Item.EndUsing == null || Item.EndUsing == DateTime.MinValue) return true;
            if (Item.EndUsing >= date.Date) return true;
            return false;
        }

        bool? _isChecked;

        /// <summary>
        /// Gets/sets the state of the associated UI toggle (ex. CheckBox).
        /// The return value is calculated based on the check state of all
        /// child FooViewModels.  Setting this property to true or false
        /// will set all children to the same check state, and setting it 
        /// to any value will cause the parent to verify its check state.
        /// </summary>
        public bool? IsChecked
        {
            get { return _isChecked; }
            set {
                CheckAllChildren((bool)value);
                this.SetIsChecked(value, true, true); }
        }

        void CheckAllChildren(bool check)
        {
            foreach (var item in Children)
            {
                (item as FooViewModel).IsChecked = check;
                //CheckAllChildren(item);
            }
        }

        CollectionViewSource _Items;
        public CollectionViewSource Items { get {
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

        public List<FooViewModel> Children
        {get; set;}

        public bool IsInitiallySelected
        {
            get
            {
                return false;
                throw new NotImplementedException();
            }
        }

        string _Name;
        public string Name
        { get { return _Name; } }

        public event PropertyChangedEventHandler PropertyChanged;

        FooViewModel _parent;
        void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (value == _isChecked)
                return;

            _isChecked = value;

            if (updateChildren && _isChecked.HasValue)
                this.Children.ForEach(c => c.SetIsChecked(_isChecked, true, false));

            if (updateParent && _parent != null)
                _parent.VerifyCheckState();

            this.OnPropertyChanged("IsChecked");
        }

        void VerifyCheckState()
        {
            bool? state = null;
            for (int i = 0; i < this.Children.Count; ++i)
            {
                bool? current = this.Children[i].IsChecked;
                if (i == 0)
                {
                    state = current;
                }
                else if (state != current)
                {
                    state = null;
                    break;
                }
            }
            this.SetIsChecked(state, false, true);
        }

        #region INotifyPropertyChanged Members

        void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
        #endregion
    }
}
