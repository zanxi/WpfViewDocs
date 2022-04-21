using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using ViewAppDocs.Model;

namespace ViewAppDocs.ViewModel
{
    internal class PersonViewModel : DependencyObject
    {

        public string FilterText
        {
            get { return (string)GetValue(FilterTextProperty); }
            set { SetValue(FilterTextProperty, value); }
        }

        // Using a DependencyProperty as the backing store for MyProperty.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty FilterTextProperty =
            DependencyProperty.Register("FilterTextProperty", typeof(string), typeof(PersonViewModel), new PropertyMetadata("", FilterText_changed));

        private static void FilterText_changed(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var current = d as PersonViewModel;
            if (current != null)
            {
                current.Items.Filter = null;
                current.Items.Filter = current.FilterPerson;
            }
        }

        public ICollectionView Items
        {
            get { return (ICollectionView)GetValue(ItemsProperty); }
            set { SetValue(ItemsProperty, value); }
        }

        public static readonly DependencyProperty ItemsProperty =
            DependencyProperty.Register("Items", typeof(ICollectionView), typeof(PersonViewModel), new PropertyMetadata(null));

        //public CollectionViewSource Items
        //{
        //    get { return (CollectionViewSource)GetValue(ItemsProperty); }
        //    set { SetValue(ItemsProperty, value); }
        //}

        //// Using a DependencyProperty as the backing store for MyProperty.  This enables animation, styling, binding, etc...
        //public static readonly DependencyProperty ItemsProperty =
        //    DependencyProperty.Register("Items", typeof(CollectionViewSource), typeof(PersonViewModel), new PropertyMetadata(null));

        public PersonViewModel()
        {
            Items = CollectionViewSource.GetDefaultView(Person.GetPersons());
            //Items = new CollectionViewSource() { Source = Person.GetPersons() };

            Items.Filter = FilterPerson;
        }

        private bool FilterPerson(object obj)
        {
            bool result = true;
            Person current = obj as Person;
            //if(current!=null&& (current.FirstName.Contains(FilterText)|| current.LastName.Contains(FilterText)))
            if (!string.IsNullOrWhiteSpace(FilterText) && current != null && !current.FirstName.Contains(FilterText) && !current.LastName.Contains(FilterText))
            {
                result = false;
            }
            return result;
        }
    }
}
