# Document Generation + Outlook Email Exel VBA

This project is a PoC showing how someone could gather data about customers, verify the fields and then generate an Invoice/Letter and then email the customer the Invoice/Letter.

Simply load your data and you will be generating letters and have the ability to send emails at high speed with little to no manual interventions.

# Getting Started?

1. Load the exel sheet in Exel ensuring that you 'Enable Content'
2. Navigate to coordinates F2 in "Sheet1" and update where the MailingTemplate can be found (you can download mine or make your own)
3. Navigate to coordinates I2 in "Sheet1" and update where you wish for the generated documents to be outputted to
4. Insert your data in to the relevant fields (C5 -> I5 inclusive)
5. Press "Data Validation" ensure that column A displays True and column B displays True (if it column B shows false the customer will not get an email)
6. Press "Generate Letters", view the 
7. Press "Send Emails", this will send the emails

# Looking Forward?

1. Add the capability for you to import data from another workbook, for common extracts such as Shopify, Stripe transactions etc.
2. Enable data sourcing from a mySQL database 

Please remember this is just a Proof of Concept

