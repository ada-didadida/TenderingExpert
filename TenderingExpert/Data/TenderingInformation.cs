using System.Collections.Generic;
using System.ComponentModel;
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

        public List<PackageInformation> LoadPackageInfo(WordReader tenderReader)
        {
            var result = new List<PackageInformation>();

            var packageContext = tenderReader.GetTableContent(1);

            int nameIndex = 0;
            int numIndex = 0;
            int budgetIndex = 0;
            int descripIndex = 0;
            int remarksIndex = 0;

            for (int i = 0; i < packageContext.Count; i++)
            {
                var row = packageContext[i];
                if (i == 0)
                {
                    for (int j = 0; j < row.Count; j++)
                    {
                        var context = row[j];

                        if (context.Contains("设备名称"))
                            nameIndex = j;

                        if (context.Contains("数量"))
                            numIndex = j;

                        if (context.Contains("预算"))
                            budgetIndex = j;

                        if (context.Contains("描述"))
                            descripIndex = j;

                        if (context.Contains("备注"))
                            remarksIndex = j;
                    }
                }
                else
                {
                    var info = new PackageInformation
                    {
                        DeviceName = row[nameIndex].Replace("\r\a", ""),
                        Quantity = row[numIndex].Replace("\r\a", ""),
                        Budget = row[budgetIndex].Replace("\r\a", ""),
                        Description = row[descripIndex].Replace("\r\a", ""),
                        Remarks = row[remarksIndex].Replace("\r\a", "")
                    };

                    result.Add(info);
                }
            }

            return result;
        }

        public List<PurchaseInformation> LoadPurchaseInfo(WordReader purchaseReader, int index)
        {
            var result = new List<PurchaseInformation>();

            var purchaseContext = purchaseReader.GetTableContent(index);

            int companyIndex = 0;
            int contactsIndex = 0;
            int mobileIndex = 0;
            int phoneIndex = 0;
            int faxIndex = 0;
            int mailIndex = 0;

            for (int i = 0; i < purchaseContext.Count; i++)
            {
                var row = purchaseContext[i];
                if (i == 0)
                {
                    for (int j = 0; j < row.Count; j++)
                    {
                        var context = row[j];

                        if (context.Contains("单位"))
                            companyIndex = j;

                        if (context.Contains("联系人"))
                            contactsIndex = j;

                        if (context.Contains("手机"))
                            mobileIndex = j;

                        if (context.Contains("电话"))
                            phoneIndex = j;

                        if (context.Contains("传真"))
                            faxIndex = j;

                        if (context.Contains("邮箱"))
                            mailIndex = j;
                    }
                }
                else
                {
                    var info = new PurchaseInformation
                    {
                        CompanyName = row[companyIndex].Replace("\r\a", ""),
                        Contacts = row[contactsIndex].Replace("\r\a", ""),
                        MobilePhone = row[mobileIndex].Replace("\r\a", ""),
                        Phone = row[phoneIndex].Replace("\r\a", ""),
                        Fax = row[faxIndex].Replace("\r\a", ""),
                        Mail = row[mailIndex].Replace("\r\a", ""),
                    };

                    result.Add(info);
                }
            }

            return result;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
