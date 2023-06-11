## Document Generation + Outlook Email Exel VBA

This project is a PoC showing how someone could gather data about customers, verify the fields and then generate an Invoice/Letter and then email the customer the Invoice/Letter.

Simply load your data and you will be generating letters and have the ability to send emails at high speed with little to no manual interventions.

## Getting Started?

1. Load the exel sheet in Exel ensuring that you 'Enable Content'
2. Navigate to coordinates F2 in "Sheet1" and update where the MailingTemplate can be found (you can download mine or make your own)
3. Navigate to coordinates I2 in "Sheet1" and update where you wish for the generated documents to be outputted to
4. Insert your data in to the relevant fields (C5 -> I5 inclusive)
5. Press "Data Validation" ensure that column A displays True and column B displays True (if it column B shows false the customer will not get an email)
6. Press "Generate Letters", view the 
7. Press "Send Emails", this will send the emails

## Looking Forward?

1. Add the capability for you to import data from another workbook, for common extracts such as Shopify, Stripe transactions etc.
2. Enable data sourcing from a mySQL database 

Please remember this is just a Proof of Concept

## Application in Action

Verifying the data has the required fields

![Alt text](https://i.gyazo.com/fa200c439ee18d6700b1c489bffdfd7c.mp4)

Following this we will generate our Letters

![Alt text](https://i.gyazo.com/c81bf28cea8ad381a06fd08ee8f98ec6.mp4)

Now that we have generated our letters we can review the output file to see if they have been generated correctly

![Alt text](https://i.gyazo.com/cb6e1e704458dfcdb4887e9b402dfb61.png)

Notice that there are now two rows in the Letter Generated table, we will now proceed to emailing the customers. This will email from the inbox you are logged in to on Outlook

![Alt text](https://i.gyazo.com/13e2b156a6969a7747a7fc840e76cdc4.mp4)

We have now been able to successfuly send our emails, and we can confirm this via our log sheet

![Alt text](https://i.gyazo.com/33476290467ce9b96246d12550dcca34.png)


