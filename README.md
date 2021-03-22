# lib_MicrosoftExcel_ui_ngx
Uses the Microsoft Excel connector for Convertigo platform. Install this library to enable reading from Microsoft Excel sheets on Office365 Cloud  for your Convertigo applications

# Installation

* In your Convertigo Studio use File->Import->Convertigo->Convertigo Project and hit the 'Next' button

* In the Dialog 'Project remote URL' field Paste :

        lib_MicrosoftExcel_ui_ngx=https://github.com/convertigo/c8oprj-lib-excel-ui-ngx.git:branch=8.0.0

* And click the 'Finish' button
* This will also automatically import the lib_OAuth project

# Usage

## Shared Actions

In order to authenticate with Microsoft Azure AD and access the availables documents in Microsoft OneDrive , the library provides a Shared Action you can use in your client apps.

Shared Action  | Usage
------| ------
Authenticate   | This will authenticate you app to Microsoft AzureAd. A Microsoft popup login page will appear to log you in. 

## Sample Application

You will find in this project a sample application using the Microsoft Excel Sheet Library, use this as a reference and tutorial about using the library. This demonstrates :
- Use of the **Authenticate** Shared Action
- use of the **SheetGetRange** Sequence

