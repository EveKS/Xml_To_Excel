using System;
using System.Xml.Serialization;
using System.Collections.Generic;
namespace Xml2CSharp
{
    [XmlRoot(ElementName = "b_n")]
    public class B_n
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "label")]
    public class Label
    {
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "b_date")]
    public class B_date
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "fo")]
    public class Fo
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "c_ca")]
    public class C_ca
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "c_mmsisdn")]
    public class C_mmsisdn
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "c_ad")]
    public class C_ad
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "c_cnum")]
    public class C_cnum
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "c_cfax")]
    public class C_cfax
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "o_ad")]
    public class O_ad
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "o_ogrn")]
    public class O_ogrn
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "o_inn")]
    public class O_inn
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_na")]
    public class Br_na
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_ad")]
    public class Br_ad
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_pm")]
    public class Br_pm
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_bn")]
    public class Br_bn
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_rs")]
    public class Br_rs
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_ks")]
    public class Br_ks
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_bik")]
    public class Br_bik
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_kpp")]
    public class Br_kpp
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
    }

    [XmlRoot(ElementName = "br_rec")]
    public class Br_rec
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "br_prim")]
    public class Br_prim
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "branch")]
    public class Branch
    {
        [XmlElement(ElementName = "br_na")]
        public Br_na Br_na { get; set; }
        [XmlElement(ElementName = "br_ad")]
        public Br_ad Br_ad { get; set; }
        [XmlElement(ElementName = "label")]
        public Label Label { get; set; }
        [XmlElement(ElementName = "br_pm")]
        public Br_pm Br_pm { get; set; }
        [XmlElement(ElementName = "br_bn")]
        public Br_bn Br_bn { get; set; }
        [XmlElement(ElementName = "br_rs")]
        public Br_rs Br_rs { get; set; }
        [XmlElement(ElementName = "br_ks")]
        public Br_ks Br_ks { get; set; }
        [XmlElement(ElementName = "br_bik")]
        public Br_bik Br_bik { get; set; }
        [XmlElement(ElementName = "br_kpp")]
        public Br_kpp Br_kpp { get; set; }
        [XmlElement(ElementName = "br_rec")]
        public Br_rec Br_rec { get; set; }
        [XmlElement(ElementName = "br_prim")]
        public Br_prim Br_prim { get; set; }
    }

    [XmlRoot(ElementName = "addr")]
    public class Addr
    {
        [XmlAttribute(AttributeName = "number")]
        public string Number { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "addrs")]
    public class Addrs
    {
        [XmlElement(ElementName = "addr")]
        public List<Addr> Addr { get; set; }
    }

    [XmlRoot(ElementName = "title")]
    public class Title
    {
        [XmlElement(ElementName = "c_name")]
        public string C_name { get; set; }
        [XmlElement(ElementName = "b_n")]
        public B_n B_n { get; set; }
        [XmlElement(ElementName = "label")]
        public List<Label> Label { get; set; }
        [XmlElement(ElementName = "b_start")]
        public string B_start { get; set; }
        [XmlElement(ElementName = "b_end")]
        public string B_end { get; set; }
        [XmlElement(ElementName = "b_date")]
        public B_date B_date { get; set; }
        [XmlElement(ElementName = "fo")]
        public Fo Fo { get; set; }
        [XmlElement(ElementName = "c_ca")]
        public C_ca C_ca { get; set; }
        [XmlElement(ElementName = "c_mmsisdn")]
        public C_mmsisdn C_mmsisdn { get; set; }
        [XmlElement(ElementName = "c_ad")]
        public C_ad C_ad { get; set; }
        [XmlElement(ElementName = "c_cnum")]
        public C_cnum C_cnum { get; set; }
        [XmlElement(ElementName = "c_cfax")]
        public C_cfax C_cfax { get; set; }
        [XmlElement(ElementName = "o_ad")]
        public O_ad O_ad { get; set; }
        [XmlElement(ElementName = "o_ogrn")]
        public O_ogrn O_ogrn { get; set; }
        [XmlElement(ElementName = "o_inn")]
        public O_inn O_inn { get; set; }
        [XmlElement(ElementName = "branch")]
        public Branch Branch { get; set; }
        [XmlElement(ElementName = "thnx")]
        public string Thnx { get; set; }
        [XmlElement(ElementName = "addrs")]
        public Addrs Addrs { get; set; }
        [XmlElement(ElementName = "tz")]
        public string Tz { get; set; }
        [XmlElement(ElementName = "print_facsimile")]
        public string Print_facsimile { get; set; }
    }

    [XmlRoot(ElementName = "c_detail")]
    public class C_detail
    {
        [XmlElement(ElementName = "c_d_n")]
        public string C_d_n { get; set; }
        [XmlElement(ElementName = "c_d_sumnd")]
        public string C_d_sumnd { get; set; }
        [XmlElement(ElementName = "c_d_sumd")]
        public string C_d_sumd { get; set; }
        [XmlElement(ElementName = "c_d_sum")]
        public string C_d_sum { get; set; }
        [XmlElement(ElementName = "c_d_id")]
        public string C_d_id { get; set; }
    }

    [XmlRoot(ElementName = "c_group")]
    public class C_group
    {
        [XmlElement(ElementName = "c_detail")]
        public C_detail C_detail { get; set; }
    }

    [XmlRoot(ElementName = "c_groups")]
    public class C_groups
    {
        [XmlElement(ElementName = "c_group")]
        public List<C_group> C_group { get; set; }
    }

    [XmlRoot(ElementName = "corr")]
    public class Corr
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "vat")]
    public class Vat
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "chrgd")]
    public class Chrgd
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "charges")]
    public class Charges
    {
        [XmlElement(ElementName = "label")]
        public List<Label> Label { get; set; }
        [XmlAttribute(AttributeName = "label")]
        public string _Label { get; set; }
        [XmlElement(ElementName = "t_sumnd")]
        public string T_sumnd { get; set; }
        [XmlElement(ElementName = "t_sumd")]
        public string T_sumd { get; set; }
        [XmlElement(ElementName = "t_sum")]
        public string T_sum { get; set; }
        [XmlElement(ElementName = "c_groups")]
        public C_groups C_groups { get; set; }
        [XmlElement(ElementName = "corr")]
        public Corr Corr { get; set; }
        [XmlElement(ElementName = "vat")]
        public Vat Vat { get; set; }
        [XmlElement(ElementName = "chrgd")]
        public Chrgd Chrgd { get; set; }
        [XmlElement(ElementName = "signa")]
        public string Signa { get; set; }
        [XmlElement(ElementName = "signb")]
        public string Signb { get; set; }
    }

    [XmlRoot(ElementName = "bal_sd")]
    public class Bal_sd
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "bal_ed")]
    public class Bal_ed
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "c_opl")]
    public class C_opl
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "payments")]
    public class Payments
    {
        [XmlElement(ElementName = "label")]
        public List<Label> Label { get; set; }
        [XmlAttribute(AttributeName = "label")]
        public string _Label { get; set; }
        [XmlElement(ElementName = "bal_sd")]
        public Bal_sd Bal_sd { get; set; }
        [XmlElement(ElementName = "bal_ss")]
        public string Bal_ss { get; set; }
        [XmlElement(ElementName = "bal_ed")]
        public Bal_ed Bal_ed { get; set; }
        [XmlElement(ElementName = "bal_es")]
        public string Bal_es { get; set; }
        [XmlElement(ElementName = "c_opl")]
        public C_opl C_opl { get; set; }
        [XmlElement(ElementName = "p_tot")]
        public string P_tot { get; set; }
        [XmlElement(ElementName = "p_table")]
        public string P_table { get; set; }
    }

    [XmlRoot(ElementName = "tot_ch")]
    public class Tot_ch
    {
        [XmlElement(ElementName = "tot_f_opu")]
        public string Tot_f_opu { get; set; }
        [XmlElement(ElementName = "tot_f_dop")]
        public string Tot_f_dop { get; set; }
        [XmlElement(ElementName = "tot_raz")]
        public string Tot_raz { get; set; }
        [XmlElement(ElementName = "tot_l_voice")]
        public string Tot_l_voice { get; set; }
        [XmlElement(ElementName = "tot_l_ic")]
        public string Tot_l_ic { get; set; }
        [XmlElement(ElementName = "tot_l_out")]
        public string Tot_l_out { get; set; }
        [XmlElement(ElementName = "tot_dop_sms")]
        public string Tot_dop_sms { get; set; }
        [XmlElement(ElementName = "tot_dop_other")]
        public string Tot_dop_other { get; set; }
        [XmlElement(ElementName = "tot_dop_pack")]
        public string Tot_dop_pack { get; set; }
        [XmlElement(ElementName = "tot_roam")]
        public string Tot_roam { get; set; }
        [XmlElement(ElementName = "tot_tot")]
        public string Tot_tot { get; set; }
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
    }

    [XmlRoot(ElementName = "total")]
    public class Total
    {
        [XmlAttribute(AttributeName = "label")]
        public string Label { get; set; }
        [XmlText]
        public string Text { get; set; }
    }

    [XmlRoot(ElementName = "charge_d")]
    public class Charge_d
    {
        [XmlElement(ElementName = "c_num")]
        public string C_num { get; set; }
        [XmlElement(ElementName = "c_rp")]
        public string C_rp { get; set; }
        [XmlElement(ElementName = "c_f_opu")]
        public string C_f_opu { get; set; }
        [XmlElement(ElementName = "c_f_dop")]
        public string C_f_dop { get; set; }
        [XmlElement(ElementName = "c_raz")]
        public string C_raz { get; set; }
        [XmlElement(ElementName = "c_l_voice")]
        public string C_l_voice { get; set; }
        [XmlElement(ElementName = "c_l_ic")]
        public string C_l_ic { get; set; }
        [XmlElement(ElementName = "c_l_out")]
        public string C_l_out { get; set; }
        [XmlElement(ElementName = "c_dop_sms")]
        public string C_dop_sms { get; set; }
        [XmlElement(ElementName = "c_dop_other")]
        public string C_dop_other { get; set; }
        [XmlElement(ElementName = "c_dop_pack")]
        public string C_dop_pack { get; set; }
        [XmlElement(ElementName = "c_roam")]
        public string C_roam { get; set; }
        [XmlElement(ElementName = "c_tot")]
        public string C_tot { get; set; }
    }

    [XmlRoot(ElementName = "charges_d")]
    public class Charges_d
    {
        [XmlElement(ElementName = "charge_d")]
        public List<Charge_d> Charge_d { get; set; }
    }

    [XmlRoot(ElementName = "ch_details")]
    public class Ch_details
    {
        [XmlElement(ElementName = "label")]
        public List<Label> Label { get; set; }
        [XmlAttribute(AttributeName = "label")]
        public string _Label { get; set; }
        [XmlElement(ElementName = "tot_ch")]
        public Tot_ch Tot_ch { get; set; }
        [XmlElement(ElementName = "total")]
        public Total Total { get; set; }
        [XmlElement(ElementName = "charges_d")]
        public Charges_d Charges_d { get; set; }
    }

    [XmlRoot(ElementName = "s_lim")]
    public class S_lim
    {
        [XmlElement(ElementName = "s_lim_item")]
        public string S_lim_item { get; set; }
    }

    [XmlRoot(ElementName = "s_d_n")]
    public class S_d_n
    {
        [XmlElement(ElementName = "str")]
        public List<string> Str { get; set; }
    }

    [XmlRoot(ElementName = "s_detail")]
    public class S_detail
    {
        [XmlElement(ElementName = "s_d_n")]
        public S_d_n S_d_n { get; set; }
        [XmlElement(ElementName = "s_d_id")]
        public string S_d_id { get; set; }
        [XmlElement(ElementName = "s_d_sum")]
        public string S_d_sum { get; set; }
        [XmlElement(ElementName = "s_d_vol")]
        public string S_d_vol { get; set; }
    }

    [XmlRoot(ElementName = "s_group")]
    public class S_group
    {
        [XmlElement(ElementName = "s_g_n")]
        public string S_g_n { get; set; }
        [XmlElement(ElementName = "s_g_id")]
        public string S_g_id { get; set; }
        [XmlElement(ElementName = "s_detail")]
        public List<S_detail> S_detail { get; set; }
    }

    [XmlRoot(ElementName = "s_groups")]
    public class S_groups
    {
        [XmlElement(ElementName = "s_group")]
        public List<S_group> S_group { get; set; }
    }

    [XmlRoot(ElementName = "sub")]
    public class Sub
    {
        [XmlElement(ElementName = "s_msisdn")]
        public string S_msisdn { get; set; }
        [XmlElement(ElementName = "s_rp")]
        public string S_rp { get; set; }
        [XmlElement(ElementName = "s_lim")]
        public List<S_lim> S_lim { get; set; }
        [XmlElement(ElementName = "s_d_tot")]
        public string S_d_tot { get; set; }
        [XmlElement(ElementName = "s_groups")]
        public S_groups S_groups { get; set; }
    }

    [XmlRoot(ElementName = "services")]
    public class Services
    {
        [XmlElement(ElementName = "sub")]
        public List<Sub> Sub { get; set; }
    }

    [XmlRoot(ElementName = "serv_det")]
    public class Serv_det
    {
        [XmlElement(ElementName = "label")]
        public List<Label> Label { get; set; }
        [XmlAttribute(AttributeName = "label")]
        public string _Label { get; set; }
        [XmlElement(ElementName = "services")]
        public Services Services { get; set; }
    }

    [XmlRoot(ElementName = "bill")]
    public class Bill
    {
        [XmlElement(ElementName = "title")]
        public Title Title { get; set; }
        [XmlElement(ElementName = "charges")]
        public Charges Charges { get; set; }
        [XmlElement(ElementName = "payments")]
        public Payments Payments { get; set; }
        [XmlElement(ElementName = "ch_details")]
        public Ch_details Ch_details { get; set; }
        [XmlElement(ElementName = "serv_det")]
        public Serv_det Serv_det { get; set; }
        [XmlAttribute(AttributeName = "id")]
        public string Id { get; set; }
    }

}
