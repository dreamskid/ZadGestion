using System.Collections.Generic;

namespace SoftwareClasses
{

    /// <summary>
    /// Type
    /// </summary>
    public enum BillType
    {
        Mission = 1,
        INVOICE = 2,
        CREDITNOTE = 3
    };

    /// <summary>
    /// Missions status
    /// </summary>
    public enum MissionStatus
    {
        NONE = 0,
        CREATED = 1,
        DONE = 2,
        DECLINED = 3,
        BILLED = 4
    };

    /// <summary>
    /// Invoice status
    /// </summary>
    public enum InvoiceStatus
    {
        NONE = 0,
        CREATED = 1,
        GENERATED = 2,
        SENT = 3,
        PAID = 4,
        UNPAID = 5
    };

    /// <summary>
    /// Credit note status
    /// </summary>
    public enum CreditNoteStatus
    {
        NONE = 0,
        CREATED = 1,
        GENERATED = 2,
        SENT = 3,
        USED = 4
    };

    /// <summary>
    /// Class Billing containing all the variables describing a bill (Mission and/or invoice)
    /// </summary>
    public class Billing
    {

        /// <summary>
        /// Copy constructor only for copy Mission
        /// </summary>
        public Billing()
        {

        }

        /// <summary>
        /// Copy constructor only for copy Mission
        /// </summary>
        public Billing(Billing _Mission)
        {
            id_HostAndHostess = _Mission.id_HostAndHostess;
            id_company = _Mission.id_company;
            amount = _Mission.amount;
            grand_amount = _Mission.grand_amount;
            discount = _Mission.discount;
            subject = _Mission.subject;
        }

        /// <summary>
        /// Getter/Setter for bill description
        /// </summary>
        public int description
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill id
        /// </summary>
        public int id
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill contact
        /// </summary>
        public string id_HostAndHostess
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill contact company
        /// </summary>
        public int id_company
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission number
        /// </summary>
        public string num_Mission
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice number
        /// </summary>
        public string num_invoice
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for credit note number
        /// </summary>
        public string num_credit
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill amount
        /// </summary>
        public double amount
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill grand amount
        /// </summary>
        public double grand_amount
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill rest amount
        /// </summary>
        public double rest_amount
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill discount
        /// </summary>
        public double discount
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill subject
        /// </summary>
        public string subject
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for credit note subject
        /// </summary>
        public string subject_credit
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission accepted date
        /// </summary>
        public string date_Mission_accepted
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission declined date
        /// </summary>
        public string date_Mission_declined
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission generated date
        /// </summary>
        public string date_Mission_generated
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice generated date
        /// </summary>
        public string date_invoice_generated
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for credit generated date
        /// </summary>
        public string date_credit_generated
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission sent date
        /// </summary>
        public string date_Mission_sent
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice sent date
        /// </summary>
        public string date_invoice_sent
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for credit sent date
        /// </summary>
        public string date_credit_sent
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice due date
        /// </summary>
        public string date_invoice_due
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice paid date
        /// </summary>
        public string date_invoice_paid
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission creation date
        /// </summary>
        public string date_Mission_creation
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice creation date
        /// </summary>
        public string date_invoice_creation
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for credit creation date
        /// </summary>
        public string date_credit_creation
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission modification date
        /// </summary>
        public string date_Mission_modification
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice modification date
        /// </summary>
        public string date_invoice_modification
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for Mission archived date
        /// </summary>
        public string date_Mission_archived
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for invoice archived date
        /// </summary>
        public string date_invoice_archived
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for credit archived date
        /// </summary>
        public string date_credit_archived
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for credit used date
        /// </summary>
        public string date_credit_used
        {
            get;
            set;
        }

        /// <summary>
        /// Getter/Setter for bill terms of payment
        /// </summary>
        public string payment_terms
        {
            get;
            set;
        }

    }
}
