using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.Sql;
using System.Data.SqlClient;
using System.Data.SqlTypes;

using System.IO;
using System.Net;
using System.Xml;
using System.Xml.Schema;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace TestXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DataSet ds;
                String sInvoiceNumber = "";
                String sPath = "";


                //fetch data for this invoice
                String connectionString = "Data Source=TOG-HOME-03;Initial Catalog=dbInvoice;Persist Security Info=True;User ID=sa;Password=w4t3rblu3;";
                String cmd = String.Format("SELECT * FROM vw_INVOICE_Final WHERE INVOICE_Final_Header_ID = {0}", 4003);
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand(cmd, conn);
                    SqlDataAdapter adapter = new SqlDataAdapter(sqlCmd);

                    ds = new DataSet("Invoice");
                    adapter.Fill(ds);
                }

                //test if datavalid
                if (ds == null)
                {
                    throw new Exception("no data");
                }
                if (ds.Tables.Count == 0)
                {
                    throw new Exception("no data");
                }
                if (ds.Tables[0].Rows.Count == 0)
                {
                    throw new Exception("no data");
                }

                //determine invoicenumber
                sInvoiceNumber = ds.Tables[0].Rows[0]["INVOICE_Final_Header_OwnInvoiceNumber"].ToString();

                //create xml document
                XmlDocument oXML = new XmlDocument();
                XmlNode oInvoice = oXML.CreateElement("Invoice");
                oXML.AppendChild(oInvoice);
                XmlAttribute oSchemaLocation = oXML.CreateAttribute("xsi", "schemaLocation", "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2 http://docs.oasis-open.org/ubl/os-UBL-2.1/xsd/maindoc/UBL-Invoice-2.1.xsd urn:www.energie-efactuur.nl:profile:invoice:ver1.0.0 ../SEeF_UBLExtension_v1.0.0.xsd");
                oSchemaLocation.Value = "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2 http://docs.oasis-open.org/ubl/os-UBL-2.1/xsd/maindoc/UBL-Invoice-2.1.xsd urn:www.energie-efactuur.nl:profile:invoice:ver1.0.0 ../SEeF_UBLExtension_v1.0.0.xsd";
                oXML.DocumentElement.SetAttributeNode(oSchemaLocation);
                oXML.DocumentElement.SetAttribute("xmlns:ext", "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2");
                oXML.DocumentElement.SetAttribute("xmlns:seef", "urn:www.energie-efactuur.nl:profile:invoice:ver1.0.0");
                oXML.DocumentElement.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                oXML.DocumentElement.SetAttribute("xmlns:cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oXML.DocumentElement.SetAttribute("xmlns:cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oXML.DocumentElement.SetAttribute("xmlns", "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2");

                //cbc:UBLVersionID
                XmlNode oUBLVersionID = oXML.CreateElement("cbc", "UBLVersionID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oUBLVersionID.InnerText = "2.1";
                oInvoice.AppendChild(oUBLVersionID);

                //cbc:CustomizationID
                XmlNode oCustomizationID = oXML.CreateElement("cbc","CustomizationID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oCustomizationID.InnerText = "urn:www.cenbii.eu:transaction:biitrns010:ver2.0: extended:urn:www.peppol.eu:bis:peppol4a:ver2.0: extended:urn:www.simplerinvoicing.org:si:si-ubl:ver1.1.x";
                oInvoice.AppendChild(oCustomizationID);

                //cbc:ProfileID
                XmlNode oProfileID = oXML.CreateElement("cbc", "ProfileID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oProfileID.InnerText = "urn:www.energie-efactuur.nl:profile:invoice:ver1.0.0";
                oInvoice.AppendChild(oProfileID);

                //cbc:ID
                XmlNode oInvoiceNumber = oXML.CreateElement("cbc", "ID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oInvoiceNumber.InnerText = sInvoiceNumber;
                oInvoice.AppendChild(oInvoiceNumber);

                //cbc:IssueDate
                XmlNode oIssueDate = oXML.CreateElement("cbc", "IssueDate", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oIssueDate.InnerText = String.Format("{0:yyyy-MM-dd}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_InvoiceDate"]);
                oInvoice.AppendChild(oIssueDate);

                //cbc:InvoiceTypeCode
                XmlNode oInvoiceTypeCode = oXML.CreateElement("cbc", "InvoiceTypeCode", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                XmlAttribute oInvoiceTypeCode_listID = oXML.CreateAttribute("listID");
                oInvoiceTypeCode_listID.Value = "UNCL1001";
                oInvoiceTypeCode.Attributes.Append(oInvoiceTypeCode_listID);
                XmlAttribute oInvoiceTypeCode_listAgencyID = oXML.CreateAttribute("listAgencyID");
                oInvoiceTypeCode_listAgencyID.Value = "6";
                oInvoiceTypeCode.Attributes.Append(oInvoiceTypeCode_listAgencyID);
                //380 debet, 384 credit
                if (Convert.ToInt32(ds.Tables[0].Rows[0]["INVOICE_Final_Header_InvoiceType_ID"]) == 1)
                {
                    oInvoiceTypeCode.InnerText = "380";
                }
                else
                {
                    oInvoiceTypeCode.InnerText = "384";
                }
                oInvoice.AppendChild(oInvoiceTypeCode);

                //cbc:Note
                XmlNode oNote = oXML.CreateElement("cbc", "Note", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oNote.InnerText = "...";
                oInvoice.AppendChild(oNote);

                //cbc:DocumentCurrencyCode -- todo:map or fetch correct value from query
                XmlNode oDocumentCurrencyCode = oXML.CreateElement("cbc", "DocumentCurrencyCode", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oDocumentCurrencyCode.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Currency_Code"]);
                oInvoice.AppendChild(oDocumentCurrencyCode);

                //cbc:AccountingCost

                //cac:InvoicePeriod
                XmlNode oInvoicePeriod = oXML.CreateElement("cac", "InvoicePeriod", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oInvoice.AppendChild(oInvoicePeriod);
                //     cbc:StartDate
                XmlNode oInvoicePeriod_StartDate = oXML.CreateElement("cbc", "StartDate", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oInvoicePeriod_StartDate.InnerText = String.Format("{0:yyyy-MM-dd}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_InvoicePeriodStart"]);
                oInvoicePeriod.AppendChild(oInvoicePeriod_StartDate);
                //     cbc:EndDate
                XmlNode oInvoicePeriod_EndDate = oXML.CreateElement("cbc", "EndDate", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oInvoicePeriod_EndDate.InnerText = String.Format("{0:yyyy-MM-dd}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_InvoicePeriodEnd"]);
                oInvoicePeriod.AppendChild(oInvoicePeriod_EndDate);

                if (String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_SalesOrderNumber"]) != "")
                {
                    //cac:OrderReference
                    XmlNode oOrderReference = oXML.CreateElement("cac", "OrderReference", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                    oInvoice.AppendChild(oOrderReference);
                    //     cbc:ID
                    XmlNode oOrderReference_ID = oXML.CreateElement("cbc", "ID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                    oOrderReference_ID.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_SalesOrderNumber"]);
                    oOrderReference.AppendChild(oOrderReference_ID);
                }

                if (String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_OwnReference"]) != "")
                {
                    //cac:BillingReference
                    XmlNode oBillingReference = oXML.CreateElement("cac", "BillingReference", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                    oInvoice.AppendChild(oBillingReference);
                    //     cac:InvoiceDocumentReference
                    XmlNode oBillingReference_InvoiceDocumentReference = oXML.CreateElement("cac", "InvoiceDocumentReference", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                    oBillingReference.AppendChild(oBillingReference_InvoiceDocumentReference);
                    //         cbc:ID
                    XmlNode oBillingReference_InvoiceDocumentReference_ID = oXML.CreateElement("cbc", "ID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                    oBillingReference_InvoiceDocumentReference_ID.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_OwnReference"]);
                    oBillingReference_InvoiceDocumentReference.AppendChild(oBillingReference_InvoiceDocumentReference_ID);
                }

                //cac: AccountingSupplyParty
                XmlNode oAccountingSupplyParty = oXML.CreateElement("cac", "AccountingSupplyParty", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oInvoice.AppendChild(oAccountingSupplyParty);
                //     cac:Party
                XmlNode oAccountingSupplyParty_Party = oXML.CreateElement("cac", "Party", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty.AppendChild(oAccountingSupplyParty_Party);
                //         cac:PartyIdentification
                XmlNode oAccountingSupplyParty_PartyIdentification = oXML.CreateElement("cac", "PartyIdentification", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty_Party.AppendChild(oAccountingSupplyParty_PartyIdentification);
                //             cbc:ID
                XmlNode oAccountingSupplyParty_PartyIdentification_ID = oXML.CreateElement("cbc", "ID", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_PartyIdentification_ID.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_KvKNumber"]);
                oAccountingSupplyParty_PartyIdentification.AppendChild(oAccountingSupplyParty_PartyIdentification_ID);
                //         cac:PartyName
                XmlNode oAccountingSupplyParty_PartyName = oXML.CreateElement("cac", "PartyName", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty_Party.AppendChild(oAccountingSupplyParty_PartyName);
                //             cbc:Name
                XmlNode oAccountingSupplyParty_PartyName_Name = oXML.CreateElement("cbc", "Name", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_PartyName_Name.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_Name"]);
                oAccountingSupplyParty_PartyName.AppendChild(oAccountingSupplyParty_PartyName_Name);
                //         cac:PostalAddress
                XmlNode oAccountingSupplyParty_Party_PostalAddress = oXML.CreateElement("cac", "PostalAddress", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty_Party.AppendChild(oAccountingSupplyParty_Party_PostalAddress);
                //             cbc:Postbox
                XmlNode oAccountingSupplyParty_Party_PostalAddress_Postbox = oXML.CreateElement("cbc", "Postbox", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_Party_PostalAddress_Postbox.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_InvoiceAddress_Line1"]);
                oAccountingSupplyParty_Party_PostalAddress.AppendChild(oAccountingSupplyParty_Party_PostalAddress_Postbox);
                //             cbc:StreetName
                XmlNode oAccountingSupplyParty_Party_PostalAddress_StreetName = oXML.CreateElement("cbc", "StreetName", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_Party_PostalAddress_StreetName.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_InvoiceAddress_Line1"]);
                oAccountingSupplyParty_Party_PostalAddress.AppendChild(oAccountingSupplyParty_Party_PostalAddress_StreetName);
                //             cbc:BuildingNumber
                XmlNode oAccountingSupplyParty_Party_PostalAddress_BuildingNumber = oXML.CreateElement("cbc", "BuildingNumber", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_Party_PostalAddress_BuildingNumber.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_InvoiceAddress_Line1"]);
                oAccountingSupplyParty_Party_PostalAddress.AppendChild(oAccountingSupplyParty_Party_PostalAddress_BuildingNumber);
                //             cbc:CityName
                XmlNode oAccountingSupplyParty_Party_PostalAddress_CityName = oXML.CreateElement("cbc", "CityName", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_Party_PostalAddress_CityName.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_InvoiceAddress_Line3"]);
                oAccountingSupplyParty_Party_PostalAddress.AppendChild(oAccountingSupplyParty_Party_PostalAddress_CityName);
                //             cbc:PostalZone
                XmlNode oAccountingSupplyParty_Party_PostalAddress_PostalZone = oXML.CreateElement("cbc", "PostalZone", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_Party_PostalAddress_PostalZone.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_InvoiceAddress_Line2"]);
                oAccountingSupplyParty_Party_PostalAddress.AppendChild(oAccountingSupplyParty_Party_PostalAddress_PostalZone);
                //             cac:Country
                XmlNode oAccountingSupplyParty_Party_PostalAddress_Country = oXML.CreateElement("cac", "Country", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty_Party_PostalAddress.AppendChild(oAccountingSupplyParty_Party_PostalAddress_Country);
                //                 cbc:IdentificationCode
                XmlNode oAccountingSupplyParty_Party_PostalAddress_Country_IdentificationCode = oXML.CreateElement("cbc", "IdentificationCode", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_Party_PostalAddress_Country_IdentificationCode.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_InvoiceAddress_Line4"]);
                oAccountingSupplyParty_Party_PostalAddress_Country.AppendChild(oAccountingSupplyParty_Party_PostalAddress_Country_IdentificationCode);
                //                 cbc:Name
                XmlNode oAccountingSupplyParty_Party_PostalAddress_Country_Name = oXML.CreateElement("cbc", "Name", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                oAccountingSupplyParty_Party_PostalAddress_Country_Name.InnerText = String.Format("{0}", ds.Tables[0].Rows[0]["INVOICE_Final_Header_Sender_InvoiceAddress_Line4"]);
                oAccountingSupplyParty_Party_PostalAddress_Country.AppendChild(oAccountingSupplyParty_Party_PostalAddress_Country_Name);
                //         cac:PartyTaxScheme
                XmlNode oAccountingSupplyParty_PartyTaxScheme = oXML.CreateElement("cac", "PartyTaxScheme", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty_Party.AppendChild(oAccountingSupplyParty_PartyTaxScheme);
                //             cbc:CompanyID
                //             cac:TaxScheme
                //                 cbc:ID
                //         cac:PartyLegalEntity
                XmlNode oAccountingSupplyParty_PartyLegalEntity = oXML.CreateElement("cac", "PartyLegalEntity", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty_Party.AppendChild(oAccountingSupplyParty_PartyLegalEntity);
                //             cbc:RegistrationName
                //             cbc:CompanyID
                //             cac:RegistrationAddress
                //                 cbc:CityName
                //                 cac:Country
                //                     cbc:IdentificationCode
                //                     cbc:Name
                //         cac:Contact
                XmlNode oAccountingSupplyParty_Contact = oXML.CreateElement("cac", "Contact", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                oAccountingSupplyParty_Party.AppendChild(oAccountingSupplyParty_Contact);
                //             cbc:Name
                //             cbc:Telephone
                //             cbc:ElectronicMail


                //determine folder / path
                sPath = String.Format(@"c:\temp\{0}.xml", sInvoiceNumber);

                //save xml to folder
                oXML.Save(sPath);

                this.richTextBox1.Text = oXML.InnerXml.Replace(">",">" + Environment.NewLine);

                //log stuff - queue etc
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
