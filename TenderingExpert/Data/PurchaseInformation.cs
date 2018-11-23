using System.ComponentModel;
using System.Runtime.CompilerServices;
using TenderingExpert.Annotations;

namespace TenderingExpert.Data
{
    public class PurchaseInformation : INotifyPropertyChanged
    {
        private string companyName;
        private string contacts;
        private string mobilePhone;
        private string phone;
        private string fax;
        private string mail;

        public string CompanyName
        {
            get => companyName;
            set
            {
                companyName = value;
                OnPropertyChanged(nameof(CompanyName));
            }
        }

        public string Contacts
        {
            get => contacts;
            set
            {
                contacts = value;
                OnPropertyChanged(nameof(Contacts));
            }
        }

        public string MobilePhone
        {
            get => mobilePhone;
            set
            {
                mobilePhone = value;
                OnPropertyChanged(nameof(MobilePhone));
            }
        }

        public string Phone
        {
            get => phone;
            set
            {
                phone = value;
                OnPropertyChanged(nameof(Phone));
            }
        }

        public string Fax
        {
            get => fax;
            set
            {
                fax = value;
                OnPropertyChanged(nameof(Fax));
            }
        }

        public string Mail
        {
            get => mail;
            set
            {
                mail = value;
                OnPropertyChanged(nameof(Mail));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
