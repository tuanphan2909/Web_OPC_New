﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace web4.Models
{
    public class CTV
    {
        public int Id { get; set; }
        public List<B30CTVDetail> Details { get; set; }

        public DateTime Ngay_Ct { get; set; }
        public string So_Ct { get; set; }
        public string Ma_Dt { get; set; }
        public string Dvcs { get; set; }
        public string Loai_TP { get; set; }
        public string Ma_vt { get; set; }
        public string Ten_Vt { get; set; }
        public string Ma_Vt_SAP { get; set; }
        public string Dvt { get; set; }
        public float Han_Muc { get; set; }
        public string Ten_Dt { get; set; }
        public string CTVId { get; set; }
        public string Ma_TDV { get; set; }
    }
}