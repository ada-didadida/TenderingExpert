using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using TenderingExpert.Annotations;
using WordOperator;

namespace TenderingExpert.Data
{
    public class TenderingInformation : INotifyPropertyChanged
    {
        private string projectName;
        private string projectCode;
        private string projectContactPerson;
        private string projectContactPhoneNumber;
        private string purchaser;
        private string purchaserAddress;
        private string purchaserContact;
        private string agency;
        private string agencyContactPerson;
        private string agencyContactPhoneNumber;
        private string agencyEmail;
        private string agencyAddress;
        private string tenderingAddress;
        private string tenderingDate;

        public string ProjectName
        {
            get => projectName;
            set
            {
                projectName = value;
                OnPropertyChanged(nameof(ProjectName));
            }
        }

        public string ProjectCode
        {
            get => projectCode;
            set
            {
                projectCode = value;
                OnPropertyChanged(nameof(ProjectCode));
            }
        }

        public string ProjectContactPerson
        {
            get => projectContactPerson;
            set
            {
                projectContactPerson = value;
                OnPropertyChanged(nameof(ProjectContactPerson));
            }
        }

        public string ProjectContactPhoneNumber
        {
            get => projectContactPhoneNumber;
            set
            {
                projectContactPhoneNumber = value;
                OnPropertyChanged(nameof(ProjectContactPhoneNumber));
            }
        }

        public string Purchaser
        {
            get => purchaser;
            set
            {
                purchaser = value;
                OnPropertyChanged(nameof(Purchaser));
            }
        }

        public string PurchaserAddress
        {
            get => purchaserAddress;
            set
            {
                purchaserAddress = value;
                OnPropertyChanged(nameof(PurchaserAddress));
            }
        }

        public string PurchaserContact
        {
            get => purchaserContact;
            set
            {
                purchaserContact = value;
                OnPropertyChanged(nameof(PurchaserContact));
            }
        }

        public string Agency
        {
            get => agency;
            set
            {
                agency = value;
                OnPropertyChanged(nameof(Agency));
            }
        }

        public string AgencyContactPerson
        {
            get => agencyContactPerson;
            set
            {
                agencyContactPerson = value;
                OnPropertyChanged(nameof(AgencyContactPerson));
            }
        }

        public string AgencyContactPhoneNumber
        {
            get => agencyContactPhoneNumber;
            set
            {
                agencyContactPhoneNumber = value;
                OnPropertyChanged(nameof(AgencyContactPhoneNumber));
            }
        }

        public string AgencyEmail
        {
            get => agencyEmail;
            set
            {
                agencyEmail = value;
                OnPropertyChanged(nameof(AgencyEmail));
            }
        }

        public string AgencyAddress
        {
            get => agencyAddress;
            set
            {
                agencyAddress = value;
                OnPropertyChanged(nameof(AgencyAddress));
            }
        }

        public string TenderingAddress
        {
            get => tenderingAddress;
            set
            {
                tenderingAddress = value;
                OnPropertyChanged(nameof(TenderingAddress));
            }
        }

        public string TenderingDate
        {
            get => tenderingDate;
            set
            {
                tenderingDate = value;
                OnPropertyChanged(nameof(TenderingDate));
            }
        }

        public void LoadInfo(WordReader reader)
        {
            ProjectName = reader.FindKeyValue("项目名称：");
            ProjectCode = reader.FindKeyValue("项目编号：");

            ProjectContactPerson = reader.FindKeyValue("项目联系人：");

            ProjectContactPhoneNumber = reader.FindKeyValue("项目联系电话：");

            Purchaser = reader.FindKeyValue("采购人：");
            PurchaserAddress = reader.FindKeyValue("地址：");
            PurchaserContact = reader.FindKeyValue("电话/传真：");

            Agency = reader.FindKeyValue("代理机构：");

            AgencyContactPerson = reader.FindKeyValue("代理机构联系人：");

            AgencyContactPhoneNumber = reader.FindKeyValue("电话：");

            AgencyEmail = reader.FindKeyValue("电子邮箱：");
            AgencyAddress = reader.FindKeyValue("代理机构地址：");

            TenderingAddress = reader.FindKeyValue("开标地址：");
            TenderingDate = reader.FindKeyValue("开标时间：");
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
