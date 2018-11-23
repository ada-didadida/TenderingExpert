using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using TenderingExpert.Annotations;

namespace TenderingExpert.Data
{
    public class PackageInformation : INotifyPropertyChanged
    {
        private string deviceName;
        private string quantity;
        private string budget;
        private string description;
        private string remarks;
        private List<PurchaseInformation> purchaseInformations;

        public string DeviceName
        {
            get => deviceName;
            set
            {
                deviceName = value;
                OnPropertyChanged(nameof(DeviceName));
            }
        }

        public string Quantity
        {
            get => quantity;
            set
            {
                quantity = value;
                OnPropertyChanged(nameof(Quantity));
            }
        }

        public string Budget
        {
            get => budget;
            set
            {
                budget = value;
                OnPropertyChanged(nameof(Budget));
            }
        }

        public string Description
        {
            get => description;
            set
            {
                description = value;
                OnPropertyChanged(nameof(Description));
            }
        }

        public string Remarks
        {
            get => remarks;
            set
            {
                remarks = value;
                OnPropertyChanged(nameof(Remarks));
            }
        }

        public List<PurchaseInformation> PurchaseInformations
        {
            get => purchaseInformations;
            set
            {
                purchaseInformations = value;
                OnPropertyChanged(nameof(PurchaseInformations));
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
