using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace OutPayslip.DataTransferObject
{
    [Serializable]
    [DataContract]
    public class ExcelDTO
    {
        [DataMember]
        public List<Sheet1> Sheet1 { get; set; }
        
    }
}