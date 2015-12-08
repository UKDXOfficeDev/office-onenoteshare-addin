# Office Addin, AngularJS and OneNote API

This is a quick sample to demonstrate authenticating and posting a page to onenote
without the user seeing a popup in the Word Desktop App. 

NB. Currently untested in Word Online, consider detecting setup and using standard popup model where necessary.  

How it works
==================

The addin has an additional element in the manifest, see `OneNoteShare.xml`

```xml
<AppDomains>
<AppDomain>https://login.live.com</AppDomain>
</AppDomains>

```
   
This allows the addin to show the login page for live without creating a popup. 

The app is then configured in the Live Application Managment portal to have a redirect URL which, following authentication, redirects back to the addins homepage. 

When the homepage is loaded the code in `shareController.js` looks for the
 `access_token` param, which is appended by live.com when redirecting. 

It extracts this and then uses it to post a new page into the users onenote. 

Getting Started
=================

First create your Live.com Application and aquire your CliendID. 

Then replace the RedirectUrl and ClientId variables at the top of `shareController.js` with your own. 

Update the `SourceLocation` element in `OneNoteShare.xml` to your hosting location and away you go. 

N.B.
========

This is a simplistic sample to show how this could be approached and should be treated as such, see comments in code for more details on limitations. 

Files of interest, /views/Home/Index.cshtml & /www/js/shareContoller, all other files are standard template. 
