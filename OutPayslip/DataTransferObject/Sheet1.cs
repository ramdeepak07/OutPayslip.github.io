using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace OutPayslip.DataTransferObject
{
    [Serializable]
    [DataContract]
    public class Sheet1
    {
        [DataMember]
        public string CompanyName { get; set; }
        [DataMember]
        public string EmployeeName { get; set; }
        [DataMember]
        public string EmpCode { get; set; }
        [DataMember]
        public Int64 InsuranceN0 { get; set; }
        [DataMember]
        public Int64 UAN { get; set; }
        [DataMember]
        public decimal FixedBasicDa { get; set; }
        [DataMember]
        public decimal FixedHRA { get; set; }
        [DataMember]
        public decimal FixedOthers { get; set; }
        [DataMember]
        public decimal FixedGross { get; set; }
        [DataMember]
        public Int64 PresentDays { get; set; }
        [DataMember]
        public decimal EarnedGrossSalary { get; set; }
        [DataMember]
        public decimal SalaryEPF { get; set; }
        [DataMember]
        public decimal SalaryESI { get; set; }
        [DataMember]
        public decimal DeductionEPf { get; set; }
        [DataMember]
        public decimal DeductionESI { get; set; }
        [DataMember]
        public decimal DeductionOthers { get; set; }
        [DataMember]
        public decimal NetSalary { get; set; }
        [DataMember]
        public decimal AdvanceGiven { get; set; }
        [DataMember]
        public decimal ActualSalary { get; set; }
    }
}