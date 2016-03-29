using ADODB;
using Hydra;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;

[assembly: CallbackAssemblyDescription
    ("Confirming",
    "Eksport zaznaczonych pozycji z preliminarza do pliku XML zgodnego ze standardem SEPA",
    "Mariusz Midzio PROSPEO",
    "1.0",
    "2016.0",
    "23-03-2016")
]

namespace Confirming
{
    [SubscribeProcedure(Procedures.Preliminarz, "Confirming")]
    public class PreliminarzPrzycisk : Callback
    {
        #region Zmienne

        public ClaWindow Przycisk, Okno, EksportXLPrzycisk;
        public int IlePrzelewow = 0;
        public decimal KwotaPrzelewow = 0;
        public DataTable ListaPlatnosci = new DataTable(), Platnosci = new DataTable();

        #endregion Zmienne

        #region Callbacki

        public override void Cleanup()
        {
        }

        public override void Init()
        {
            Okno = GetWindow();
            AddSubscription(true, 0, Events.OpenWindow, new TakeEventDelegate(OpenWindow));
            AddSubscription(true, 0, Events.ResizeWindow, new TakeEventDelegate(ResizeWindow));
            AddSubscription(true, 0, Events.Maximize, new TakeEventDelegate(ResizeWindow));
            AddSubscription(true, 0, Events.Maximize, new TakeEventDelegate(ResizeWindow));
        }

        private bool OpenWindow(Procedures ProcID, int ControlID, Events Event)
        {
            Przycisk = Okno.AllChildren["?PreliminarzTab"].AllChildren.Add(ControlTypes.button);
            Przycisk.TextRaw = "CF";
            Przycisk.ToolTipRaw = "Eksportuje zaznaczone płatności do pliku w standardzie SEPA";
            Przycisk.FontStyleRaw = "700";
            Przycisk.Enabled = true;
            Przycisk.Visible = true;
            AddSubscription(false, Przycisk.Id, Events.Accepted, new TakeEventDelegate(GenerujPlik));
            //Podpięcie po ikonę zwijającą filtr
            AddSubscription(false, Okno.AllChildren["?bUkryjFiltrDodatkowy"].Id, Events.Accepted, new TakeEventDelegate(ResizeWindow));

            OdswiezKontrolki();

            return true;
        }

        private bool ResizeWindow(Procedures ProcID, int ControlID, Events Event)
        {
            OdswiezKontrolki();
            return true;
        }

        private void OdswiezKontrolki()
        {
            EksportXLPrzycisk = Okno.AllChildren["?EksportButton"];
            Rectangle PolozeniePrzyciskuXL = EksportXLPrzycisk.Bounds;
            Przycisk.Bounds = new Rectangle(PolozeniePrzyciskuXL.X - 20, PolozeniePrzyciskuXL.Y, PolozeniePrzyciskuXL.Width, PolozeniePrzyciskuXL.Height);
        }

        #endregion Callbacki

        #region Obsługa

        private bool GenerujPlik(Procedures ProcID, int ControlID, Events Event)
        {
            PobierzListeZaznaczonych(ProcID);
            PobierzSzczegolyPlatnosci();
            IlePrzelewow = ListaPlatnosci.Rows.Count;
            KwotaPrzelewow = Convert.ToDecimal(ListaPlatnosci.Compute("SUM(Kwota)","1=1"));
            GenerujXML();
            Runtime.WindowController.LockThread();
            if (MessageBox.Show("Poprawnie wygenerowano plik." + Environment.NewLine + "Czy chcesz zmienić status przelewów na <wysłane>?", "Pytanie", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                ZmienStatusPlatnosci();
            }
            Runtime.WindowController.UnlockThread();
            Runtime.WindowController.PostEvent(Okno.AllChildren["?PreliminarzBrowse"].Id, Events.FullRefresh);
            return true;
        }

        private void PobierzListeZaznaczonych(Procedures ProcID)
        {
            // Dodaj zaznaczenia do tabeli
            Platnosci.Clear();
            Platnosci.Columns.Clear();
            Platnosci.Columns.Add("GIDNUmer", typeof(int));
            Platnosci.Columns.Add("GIDTyp", typeof(int));
            Platnosci.Columns.Add("GIDLp", typeof(int));

            string fieldName;
            int GIDNumer = 0, GIDTyp = 0, GIDLp = 0;
            int listaId = Okno.AllChildren["?PreliminarzBrowse"].Id;

            _Recordset recordset = Runtime.WindowController.GetQueueMarked((int)ProcID, listaId, GetCallbackThread());
            try
            {
                //jesli nie jest nic zaznaczone to recordset == null
                if (recordset != null && recordset.RecordCount > 0)
                {
                    recordset.MoveFirst();

                    while (recordset.EOF == false)
                    {
                        ADODB.Fields fields = recordset.Fields;

                        for (int i = 0; i < fields.Count; i++)
                        {
                            fieldName = fields[i].Name;
                            if (fieldName == "NUMER")
                            {
                                GIDNumer = Convert.ToInt32(fields[i].Value);
                            }
                            if (fieldName == "TYP")
                            {
                                GIDTyp = Convert.ToInt32(fields[i].Value);
                            }
                            if (fieldName == "LP")
                            {
                                GIDLp = Convert.ToInt32(fields[i].Value);
                            }
                        }
                        Platnosci.Rows.Add(GIDNumer, GIDTyp, GIDLp);

                        recordset.MoveNext();
                    }
                }
                IlePrzelewow = recordset.RecordCount;
            }
            catch (Exception ex)
            {
                Runtime.WindowController.LockThread();
                MessageBox.Show(ex.ToString(), "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                Runtime.WindowController.UnlockThread();
            }
        }

        private void GenerujXML()
        {
            SaveFileDialog ZapiszPlikDialog = new SaveFileDialog();
            ZapiszPlikDialog.Title = "Zapisz plik eksportu";
            ZapiszPlikDialog.Filter = "Pliki XML|*.xml";
            if (ZapiszPlikDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream PlikWynikowy = new FileStream(ZapiszPlikDialog.FileName, FileMode.Create);
                //Towrzę plik XML
                //Nagłówek
                XmlTextWriter XMLPlik = new XmlTextWriter(PlikWynikowy, Encoding.UTF8);
                //Root dokumentu
                XMLPlik.WriteStartDocument();
                XMLPlik.WriteStartElement("Document");
                XMLPlik.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                XMLPlik.WriteAttributeString("xmlns", "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03");
                XMLPlik.WriteStartElement("CstmrCdtTrfInitn");

                #region Nagłowek dokumentu XML

                XMLPlik.WriteStartElement("GrpHdr");
                //MsgId
                XMLPlik.WriteStartElement("MsgId");
                XMLPlik.WriteString("00/09/2015001"); //wartosć stała w każdej paczce
                XMLPlik.WriteEndElement();
                //CreDtTm
                XMLPlik.WriteStartElement("CreDtTm");
                XMLPlik.WriteString(DateTime.Today.ToString()); //data wygenerowania pliku z systemu
                XMLPlik.WriteEndElement();
                //NbOfTxs
                XMLPlik.WriteStartElement("NbOfTxs");
                XMLPlik.WriteString(IlePrzelewow.ToString()); //liczba przelewów w paczce
                XMLPlik.WriteEndElement();
                //CtrlSum
                XMLPlik.WriteStartElement("CtrlSum");
                XMLPlik.WriteString(KwotaPrzelewow.ToString().Replace(",",".")); //wartość wszystkich przelewów
                XMLPlik.WriteEndElement();
                //InitgPty
                XMLPlik.WriteStartElement("InitgPty"); //Zlecającu confirming
                //Nm
                XMLPlik.WriteStartElement("Nm");
                XMLPlik.WriteString(Nazwa());
                XMLPlik.WriteEndElement();
                //Id
                XMLPlik.WriteStartElement("Id");
                //OrgId
                XMLPlik.WriteStartElement("OrgId");
                //Othr
                XMLPlik.WriteStartElement("Othr");
                //Id
                XMLPlik.WriteStartElement("Id");
                XMLPlik.WriteString("00/09/2015001"); //wartosć stała w każdej paczce
                XMLPlik.WriteEndElement();
                //Koniec Othr
                XMLPlik.WriteEndElement();
                //Koniec OrgId
                XMLPlik.WriteEndElement();
                //Koniec ID
                XMLPlik.WriteEndElement();
                //Koniec InitgPty
                XMLPlik.WriteEndElement();
                //Koniec GrpHdr
                XMLPlik.WriteEndElement();

                #endregion Nagłowek dokumentu XML

                #region Sekcja elementów

                XMLPlik.WriteStartElement("PmtInf");

                #region Elemty Podsumowanie

                //PmtInfId
                XMLPlik.WriteStartElement("PmtInfId");
                XMLPlik.WriteString("00/09/2015001"); //stała wartość
                XMLPlik.WriteEndElement();
                //PmtMtd
                XMLPlik.WriteStartElement("PmtMtd");
                XMLPlik.WriteString("TRF"); //Stała wartosść
                XMLPlik.WriteEndElement();
                //BtchBookg
                XMLPlik.WriteStartElement("BtchBookg");
                XMLPlik.WriteString("true");    //Stała wartość
                XMLPlik.WriteEndElement();
                //NbOfTxs
                XMLPlik.WriteStartElement("NbOfTxs");
                XMLPlik.WriteString(IlePrzelewow.ToString()); //liczba wszystkich przelewów
                XMLPlik.WriteEndElement();
                //CtrlSum
                XMLPlik.WriteStartElement("CtrlSum");
                XMLPlik.WriteString(KwotaPrzelewow.ToString().Replace(",", ".")); //wartosć wszystkich przelewów
                XMLPlik.WriteEndElement();
                //ReqdExctnDt
                XMLPlik.WriteStartElement("ReqdExctnDt");
                XMLPlik.WriteString(DateTime.Today.ToString());
                XMLPlik.WriteEndElement();

                #region Dane zlecającego

                //Dbtr
                XMLPlik.WriteStartElement("Dbtr");
                //Nm
                XMLPlik.WriteStartElement("Nm");
                XMLPlik.WriteString(Nazwa());
                XMLPlik.WriteEndElement();
                //Id
                XMLPlik.WriteStartElement("Id");
                //OrgId
                XMLPlik.WriteStartElement("OrgId");
                //Othr
                XMLPlik.WriteStartElement("Othr");
                //Id
                XMLPlik.WriteStartElement("Id");
                XMLPlik.WriteString("00/09/2015001"); //wartosć stała w każdej paczce
                XMLPlik.WriteEndElement();
                //Koniec Othr
                XMLPlik.WriteEndElement();
                //Koniec OrgId
                XMLPlik.WriteEndElement();
                //Koniec ID
                XMLPlik.WriteEndElement();

                //Koniec Dbtr
                XMLPlik.WriteEndElement();

                #endregion Dane zlecającego

                #region Dane zlecającego konto

                //DbtrAcct
                XMLPlik.WriteStartElement("DbtrAcct");
                //Id
                XMLPlik.WriteStartElement("Id");
                //IBAN
                XMLPlik.WriteStartElement("IBAN");
                XMLPlik.WriteString(NrRachunku()); //nasz numer IBAN -musi być w konfiguracji IBAN
                XMLPlik.WriteEndElement();
                //Koniec Id
                XMLPlik.WriteEndElement();
                //Ccy
                XMLPlik.WriteStartElement("Ccy");
                XMLPlik.WriteString(Runtime.ConfigurationDictionary.WalutaSystemowa); //waluta
                XMLPlik.WriteEndElement();
                //Koniec DbtrAcct
                XMLPlik.WriteEndElement();

                #endregion Dane zlecającego konto

                #region Dane zlecajacego numer BIC

                //DbtrAgt
                XMLPlik.WriteStartElement("DbtrAgt");
                //FinInstnId
                XMLPlik.WriteStartElement("FinInstnId");
                //IBAN
                XMLPlik.WriteStartElement("BIC");
                XMLPlik.WriteString(BIC()); //numer BIC - wartość stała
                XMLPlik.WriteEndElement();
                //Koniec FinInstnId
                XMLPlik.WriteEndElement();
                //Koniec DbtrAgt
                XMLPlik.WriteEndElement();

                #endregion Dane zlecajacego numer BIC

                #endregion Elemty Podsumowanie

                #region Elementy

                //Pętla po zaznaczonych płatnościach
                foreach (DataRow platnosc in ListaPlatnosci.Rows)
                {
                    #region CdtTrfTxInf

                    XMLPlik.WriteStartElement("CdtTrfTxInf");

                    #region PmtId

                    XMLPlik.WriteStartElement("PmtId");
                    //EndToEndId
                    XMLPlik.WriteStartElement("EndToEndId");
                    XMLPlik.WriteString(platnosc[9].ToString());
                    XMLPlik.WriteEndElement();
                    //Koniec PmtId
                    XMLPlik.WriteEndElement();

                    #endregion PmtId

                    #region Amt

                    XMLPlik.WriteStartElement("Amt");
                    //InstdAmt
                    XMLPlik.WriteStartElement("InstdAmt");
                    XMLPlik.WriteAttributeString("Ccy", "PLN");
                    XMLPlik.WriteString(platnosc[7].ToString().Replace(",","."));
                    XMLPlik.WriteEndElement();
                    //koniec Amt
                    XMLPlik.WriteEndElement();

                    #endregion Amt

                    #region CdtrAgt

                    //CdtrAgt
                    XMLPlik.WriteStartElement("CdtrAgt");
                    //FinInstnId
                    XMLPlik.WriteStartElement("FinInstnId");
                    //BIC
                    XMLPlik.WriteStartElement("BIC");
                    XMLPlik.WriteString(platnosc[6].ToString());
                    XMLPlik.WriteEndElement();
                    //Koniec FinInstnId
                    XMLPlik.WriteEndElement();
                    //koniec CdtrAgt
                    XMLPlik.WriteEndElement();

                    #endregion CdtrAgt

                    #region Cdtr

                    XMLPlik.WriteStartElement("Cdtr");
                    //Nm
                    XMLPlik.WriteStartElement("Nme");
                    XMLPlik.WriteString(platnosc[0].ToString());
                    XMLPlik.WriteEndElement();
                    //PstlAdr
                    XMLPlik.WriteStartElement("PstlAdr");
                    //Ctry
                    XMLPlik.WriteStartElement("Ctry");
                    XMLPlik.WriteString(platnosc[1].ToString());
                    XMLPlik.WriteEndElement();
                    //AdrLine
                    XMLPlik.WriteStartElement("Ctry");
                    XMLPlik.WriteString(platnosc[10].ToString() + ", " + platnosc[2].ToString() + ".  " + platnosc[3].ToString() + "  " + platnosc[11].ToString());
                    XMLPlik.WriteEndElement();
                    //Id
                    XMLPlik.WriteStartElement("Id");
                    //PrvtId
                    XMLPlik.WriteStartElement("PrvtId");
                    //Id
                    XMLPlik.WriteStartElement("Id");
                    XMLPlik.WriteString(platnosc[4].ToString()); //NIP
                    XMLPlik.WriteEndElement();

                    //Koniec PrvtId
                    XMLPlik.WriteEndElement();
                    //Koniec ID
                    XMLPlik.WriteEndElement();

                    //Koniec PstlAdr
                    XMLPlik.WriteEndElement();

                    //koniec Cdtr
                    XMLPlik.WriteEndElement();

                    #endregion Cdtr

                    #region CdtrAcct

                    //CdtrAcct
                    XMLPlik.WriteStartElement("CdtrAcct");
                    //Id
                    XMLPlik.WriteStartElement("Id");
                    //IBAN
                    XMLPlik.WriteStartElement("IBAN");
                    XMLPlik.WriteString("PL"+platnosc[5].ToString().Replace(" ","").Replace("-","")); //IBAN
                    XMLPlik.WriteEndElement();

                    //Koniec Id
                    XMLPlik.WriteEndElement();
                    //Koniec CdtrAcct
                    XMLPlik.WriteEndElement();

                    #endregion CdtrAcct

                    //Koniec CdtTrfTxInf
                    XMLPlik.WriteEndElement();

                    #endregion CdtTrfTxInf
                }

                #endregion Elementy

                //Koniec PmtInf
                XMLPlik.WriteEndElement();

                #endregion Sekcja elementów

                //Koniec CstmrCdtTrfInitn
                XMLPlik.WriteEndElement();
                //Koniec Document
                XMLPlik.WriteEndElement();
                XMLPlik.Flush();
                PlikWynikowy.Close();
            }
        }

        private void PobierzSzczegolyPlatnosci()
        {
            ListaPlatnosci.Clear();
            ListaPlatnosci.Columns.Clear();

            ListaPlatnosci.Columns.Add("Nazwa", typeof(string));
            ListaPlatnosci.Columns.Add("Kraj", typeof(string));
            ListaPlatnosci.Columns.Add("KodPocztowy", typeof(string));
            ListaPlatnosci.Columns.Add("Miasto", typeof(string));
            ListaPlatnosci.Columns.Add("Nip", typeof(string));
            ListaPlatnosci.Columns.Add("NrRachunku", typeof(string));
            ListaPlatnosci.Columns.Add("SWIFT", typeof(string));
            ListaPlatnosci.Columns.Add("Kwota", typeof(decimal));
            ListaPlatnosci.Columns.Add("Waluta", typeof(string));
            ListaPlatnosci.Columns.Add("NrDok", typeof(string));
            ListaPlatnosci.Columns.Add("Adres", typeof(string));
            ListaPlatnosci.Columns.Add("Powiat", typeof(string));

            foreach (DataRow wiersz in Platnosci.Rows)
            {
                string sql = "SELECT\n"
           + "	ka.KnA_Nazwa1,\n"
           + "	ka.KnA_Kraj,\n"
           + "	REPLACE(ka.KnA_KodP,'-',''),\n"
           + "	ka.KnA_Miasto,\n"
           + "	ka.KnA_Nip,\n"
           + "	KnA_NrRachunku,\n"
           + "	b.Bnk_Swift,\n"
           + "	tp.TrP_Pozostaje,\n"
           + "	tp.TrP_Waluta,\n"
           + "	REPLACE(tn.TrN_DokumentObcy,'/',''),\n"
           + "	ka.KnA_Ulica,\n"
           + "	ka.KnA_Powiat\n"
           + "FROM CDN.TraPlat tp\n"
           + "JOIN CDN.KntAdresy ka ON tp.TrP_KnANumer = ka.KnA_GIDNumer\n"
           + "JOIN CDN.TraNag tn ON tn.TrN_GIDTyp = tp.TrP_GIDTyp AND tn.TrN_GIDNumer = tp.TrP_GIDNumer\n"
           + "LEFT JOIN CDN.Banki b ON b.Bnk_GIDNumer = ka.KnA_BnkNumer\n"
           + "WHERE TrP_GIDTyp={1} AND TrP_GIDNumer={0} AND TrP_GIDLp={2}";

                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = string.Format(sql, wiersz[0].ToString(), wiersz[1].ToString(), wiersz[2].ToString());
                cmd.Connection = Runtime.ActiveRuntime.Repository.Connection;
                if (Runtime.ActiveRuntime.Repository.Connection.State == ConnectionState.Closed)
                {
                    cmd.Connection.Open();
                }
                SqlDataReader Dane = cmd.ExecuteReader();
                while (Dane.Read())
                {
                    ListaPlatnosci.Rows.Add
                        (
                        Dane[0].ToString(),
                        Dane[1].ToString(),
                        Dane[2].ToString(),
                        Dane[3].ToString(),
                        Dane[4].ToString(),
                        Dane[5].ToString(),
                        Dane[6].ToString(),
                        Dane[7].ToString(),
                        Dane[8].ToString(),
                        Dane[9].ToString(),
                        Dane[10].ToString(),
                        Dane[11].ToString()
                        );
                }
                cmd.Connection.Close();
            }
        }

        private void ZmienStatusPlatnosci()
        {
            string sql = "UPDATE CDN.TraPlat\n"
           + "SET\n"
           + "CDN.TraPlat.TrP_Status = 2\n"
           + "WHERE CDN.TraPlat.TrP_GIDTyp = {1} AND CDN.TraPlat.TrP_GIDNumer = {0} AND CDN.TraPlat.TrP_GIDLp = {2}";
            foreach (DataRow wiersz in Platnosci.Rows)
            {
                Runtime.Config.ExecSql(string.Format(sql, wiersz[0].ToString(), wiersz[1].ToString(), wiersz[2].ToString()), true);
            }
        }

        private string Nazwa()
        {
            string Wynik = string.Empty;
            string sql = "SELECT RTRIM(f.Frm_Nazwa1+' '+f.Frm_Nazwa2) FROM CDN.Firma f WHERE f.Frm_GidLp = {0}";

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(sql, Runtime.Config.PieczatkaNazwa);
            cmd.Connection = Runtime.ActiveRuntime.Repository.Connection;
            if (Runtime.ActiveRuntime.Repository.Connection.State == System.Data.ConnectionState.Closed)
            {
                cmd.Connection.Open();
            }
            Wynik = (string)cmd.ExecuteScalar();
            cmd.Connection.Close();

            return Wynik;
        }

        private string NrRachunku()
        {
            string Wynik = "PL2483220004110000002367";
            //string sql = "SELECT RTRIM(f.Frm_Nazwa1+' '+f.Frm_Nazwa2) FROM CDN.Firma f WHERE f.Frm_GidLp = {0}";

            //SqlCommand cmd = new SqlCommand();
            //cmd.CommandText = string.Format(sql, Runtime.Config.PieczatkaNazwa);
            //cmd.Connection = Runtime.ActiveRuntime.Repository.Connection;
            //if (Runtime.ActiveRuntime.Repository.Connection.State == System.Data.ConnectionState.Closed)
            //{
            //    cmd.Connection.Open();
            //}
            //Wynik = (string)cmd.ExecuteScalar();
            //cmd.Connection.Close();

            return Wynik;
        }

        private string BIC()
        {
            string Wynik = "CAIXPLPWXXX";
            //string sql = "SELECT RTRIM(f.Frm_Nazwa1+' '+f.Frm_Nazwa2) FROM CDN.Firma f WHERE f.Frm_GidLp = {0}";

            //SqlCommand cmd = new SqlCommand();
            //cmd.CommandText = string.Format(sql, Runtime.Config.PieczatkaNazwa);
            //cmd.Connection = Runtime.ActiveRuntime.Repository.Connection;
            //if (Runtime.ActiveRuntime.Repository.Connection.State == System.Data.ConnectionState.Closed)
            //{
            //    cmd.Connection.Open();
            //}
            //Wynik = (string)cmd.ExecuteScalar();
            //cmd.Connection.Close();

            return Wynik;
        }

        #endregion Obsługa
    }
}