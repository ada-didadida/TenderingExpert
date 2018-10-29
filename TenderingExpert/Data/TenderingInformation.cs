using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using WordOperator;

namespace TenderingExpert.Data
{
    class TenderingInformation
    {
        public string ProjectName { get; set; }

        public string ProjectCode { get; set; }

        public List<string> ProjectContactPerson { get; set; }

        public List<string> ProjectContactPhoneNumber { get; set; }

        public string Purchaser { get; set; }

        public string PurchaserAddress { get; set; }

        public string PurchaserContact { get; set; }

        public string Agency { get; set; }

        public List<string> AgencyContactPerson { get; set; }

        public List<string> AgencyContactPhoneNumber { get; set; }

        public string AgencyEmail { get; set; }

        public string AgencyAddress { get; set; }

        public TenderingInformation(WordReader reader)
        {
            ProjectName = reader.FindKeyValue("项目名称：");
            ProjectCode = reader.FindKeyValue("项目编号：");

            var persons = reader.FindKeyValue("项目联系人：");
            ProjectContactPerson = persons.Split('，').ToList();

            var phoneNumber = reader.FindKeyValue("项目联系电话：");
            ProjectContactPhoneNumber = phoneNumber.Split('，').ToList();

            Purchaser = reader.FindKeyValue("采购人：");
            PurchaserAddress = reader.FindKeyValue("地址：");
            PurchaserContact = reader.FindKeyValue("电话/传真：");

            Agency = reader.FindKeyValue("代理机构：");

            var agencyPersons = reader.FindKeyValue("代理机构联系人：");
            AgencyContactPerson = agencyPersons.Split('，').ToList();

            var agencyPhoneNumber = reader.FindKeyValue("电话：");
            AgencyContactPhoneNumber = agencyPhoneNumber.Split('，').ToList();

            AgencyEmail = reader.FindKeyValue("电子邮箱：");
            AgencyAddress = reader.FindKeyValue("代理机构地址：");
        }
    }
}
