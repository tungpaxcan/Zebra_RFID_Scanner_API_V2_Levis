//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Zebra_RFID_Scanner.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class General
    {
        public int Id { get; set; }
        public string IdReports { get; set; }
        public string So { get; set; }
        public string Po { get; set; }
        public string Sku { get; set; }
        public string CartonTo { get; set; }
        public string UPC { get; set; }
        public string Qty { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public string CreateBy { get; set; }
        public Nullable<System.DateTime> ModifyDate { get; set; }
        public string ModifyBy { get; set; }
        public Nullable<bool> Status { get; set; }
        public string cntry { get; set; }
        public string port { get; set; }
        public string deviceNum { get; set; }
        public string doNo { get; set; }
        public string setCd { get; set; }
        public string subDoNo { get; set; }
        public string mngFctryCd { get; set; }
        public string facBranchCd { get; set; }
        public string packKey { get; set; }
        public string Location { get; set; }
        public string LocationClient { get; set; }
        public Nullable<bool> StatusClient { get; set; }
        public string deviceNumClient { get; set; }
    
        public virtual Report Report { get; set; }
    }
}
