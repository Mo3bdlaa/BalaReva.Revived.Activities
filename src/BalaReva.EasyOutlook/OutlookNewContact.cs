namespace BalaReva.EasyOutlook;

/// <summary>
/// The details of a contact to add to Outlook.
/// </summary>
/// <remarks>
/// Passed to <c>NewContact</c>. Members mirror the matching
/// <c>Microsoft.Office.Interop.Outlook.ContactItem</c> properties; anything left unset
/// is left unset on the created contact.
/// </remarks>
public sealed class OutlookNewContact
{
    /// <summary>Contact's birthday.</summary>
    public DateTime Birthday { get; set; }

    /// <summary>Contact's business 2 telephone number.</summary>
    public string Business2TelephoneNumber { get; set; } = string.Empty;

    /// <summary>Contact's business address.</summary>
    public string BusinessAddress { get; set; } = string.Empty;

    /// <summary>Contact's business address city.</summary>
    public string BusinessAddressCity { get; set; } = string.Empty;

    /// <summary>Contact's business address country.</summary>
    public string BusinessAddressCountry { get; set; } = string.Empty;

    /// <summary>Contact's business address post office box.</summary>
    public string BusinessAddressPostOfficeBox { get; set; } = string.Empty;

    /// <summary>Contact's business address postal code.</summary>
    public string BusinessAddressPostalCode { get; set; } = string.Empty;

    /// <summary>Contact's business address state.</summary>
    public string BusinessAddressState { get; set; } = string.Empty;

    /// <summary>Contact's business address street.</summary>
    public string BusinessAddressStreet { get; set; } = string.Empty;

    /// <summary>Companies associated with the contact.</summary>
    public string Companies { get; set; } = string.Empty;

    /// <summary>Main telephone number for the contact's company.</summary>
    public string CompanyMainTelephoneNumber { get; set; } = string.Empty;

    /// <summary>Contact's company name.</summary>
    public string CompanyName { get; set; } = string.Empty;

    /// <summary>Contact's email 1 address.</summary>
    public string Email1Address { get; set; } = string.Empty;

    /// <summary>Contact's email 2 address.</summary>
    public string Email2Address { get; set; } = string.Empty;

    /// <summary>Contact's first name.</summary>
    public string FirstName { get; set; } = string.Empty;

    /// <summary>Contact's home 2 telephone number.</summary>
    public string Home2TelephoneNumber { get; set; } = string.Empty;

    /// <summary>Contact's home address.</summary>
    public string HomeAddress { get; set; } = string.Empty;

    /// <summary>Contact's home address city.</summary>
    public string HomeAddressCity { get; set; } = string.Empty;

    /// <summary>Contact's home address country.</summary>
    public string HomeAddressCountry { get; set; } = string.Empty;

    /// <summary>Contact's home address post office box.</summary>
    public string HomeAddressPostOfficeBox { get; set; } = string.Empty;

    /// <summary>Contact's home address postal code.</summary>
    public string HomeAddressPostalCode { get; set; } = string.Empty;

    /// <summary>Contact's home address state.</summary>
    public string HomeAddressState { get; set; } = string.Empty;

    /// <summary>Contact's home address street.</summary>
    public string HomeAddressStreet { get; set; } = string.Empty;

    /// <summary>Contact's home fax number.</summary>
    public string HomeFaxNumber { get; set; } = string.Empty;

    /// <summary>Contact's home telephone number.</summary>
    public string HomeTelephoneNumber { get; set; } = string.Empty;

    /// <summary>Contact's last name.</summary>
    public string LastName { get; set; } = string.Empty;

    /// <summary>Contact's middle name.</summary>
    public string MiddleName { get; set; } = string.Empty;

    /// <summary>Contact's mobile telephone number.</summary>
    public string MobileTelephoneNumber { get; set; } = string.Empty;

    /// <summary>Contact's nick name.</summary>
    public string NickName { get; set; } = string.Empty;
}
