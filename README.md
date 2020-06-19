# CNDAtools
VSTO Add-in tools for generating and emailing powerpoint presentations requiring a watermark.  This is a Visual Studio solution containing two main add-ins.  CNDAEmail is the solution.  Add-ins are CNDAPowerpoint and CNDAOutlook.  CNDAExcel is deprecated.

## CNDAPowerpoint
Powerpoint VSTO add-in to substitute a list of strings with values contained in a master XML file and generate a PDF document.
Currently used to help watermark confidential documents

### XML Input
```
<?xml version="1.0" encoding="UTF-8"?>
<ArrayOfCustomer xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <customer cnda="12345" name="ACME Anvils">
    <edit key="CNDA#+">12345</edit>
    <edit key="CustName">ACME Anvils</edit>
  </customer>
  <customer cnda="11111" name="Sanford and Sons">
    <edit key="CNDA#+">11111</edit>
    <edit key="CustName">Sanford & Sons</edit>
  </customer>
</ArrayOfCustomer>
```
Where 
* **cnda** is the cnda number used in the export filename.
* **name** is the name used in the export filename.
* **key** is a regular expression to search and replace with the **edit** value

### Output
The PDF filename will be a concatination of the PowerPoint filename + the **cnda** number + the **name**

## CNDAOutlook
Outlook VSTO add-in to generate emails based on the current template email to each customer listed in the XML file.  The generated email we default into the **Drafts** folder of Outlook, though the add-in has the capability to change that location.  One version will take as input a PowerPoint file that was watermarked and PDF'ed as above and attach it to the new emails with the correct customer.
### XML Input
```
<?xml version="1.0" encoding="UTF-8"?>
<ArrayOfCustomer xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <customer cnda="12345" name="ACME Anvils">
    <edit key="CNDA#+">12345</edit>
    <edit key="CustName">ACME Anvils</edit>
    <address type="MailTo">Roadrunner@acme.com</address>
    <address type="MailTo">Coyote@acme.com</address>
    <address type="MailCC">Bugs@acme.com</address>
    <address type="MailBCC">Elmer@acme.com</address>
  </customer>
  <customer cnda="11111" name="Sanford and Sons">
    <edit key="CNDA#+">11111</edit>
    <edit key="CustName">Sanford & Sons</edit>
    <address type="MailTo">Lamont@sns.com</address>
    <address type="MailCC">Fred@sns.com</address>
    <address type="MailBCC">Esther@sns.com</address>
  </customer>
</ArrayOfCustomer>
```
Where
* **address** is an individual email address
* **type** is where that email will reside in the message.  Can only be one of three below.
  * **MailTo** this address will show up on the To: line
  * **MailCC** this address will show up on the CC: line
  * **MailBCC** this address will show up on the BCC: line

### _CNDAExcel_
_Deprecated - Initially used Excel to contain the data, but moved to XML._
