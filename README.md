# gs-to-harvestapp

Integrates Google Sheets with Harvest App.

This script requires the use of an authentication library.
See OAuth2 library documentation at: https://github.com/googlesamples/apps-script-oauth2

This implementation is proprietary to some sheets that I am using, but the core functionality for integrating with Google Sheets (such as the authentication) is still here.

These are the instructions which exist inside the "Controls" sheet, and are what allows any user in any project to utilize this script:

```
1. Enter the name of the sheet which you want to grab from into B7. (must be EXACT))
2. Log into Harvest.		https://YOURACCOUNT.harvestapp.com/
3. In Harvest, navigate to the project that you wish to use. Copy the last set of numbers in its URL into B16. This is the project's ID.	4. Visit your Harvest project's OAuth2 clients list using the URL in C17.		https://platform.harvestapp.com/oauth2_clients/
5. Create a new client in Harvest's OAuth2 clients list. Choose a meaningful App Name and Website.
6. Run "Harvest>Generate redirect URI" from the dropdown menu.
7. Copy the URI in B 20 to the "Redirect URI" section in Harvest.
8. Press the save button.
9. The page should have refreshed, and there should be a Client Parameters section at the bottom.
10. Copy the Client ID into B23.
11. Copy the Client Secret into B24.
```

I'm aware that this isn't entirely insightful documentation, but it's a start and it's better than nothing.

---

Some notes on OAuth2 and Basic Authentication:

Visit https://platform.harvestapp.com/oauth2_clients/ to set up an OAuth2 client.
For the redirect URI required in Harvest's site:
- Retrieve Google script project code from File>Project Properties>Info>Project Key
- Use that code in the redirect URI which looks like:
- https://script.google.com/macros/d/PROJECT CODE HERE/usercallback
When setting up the authentication, this message is sent via the authentication library:
- https://example.harvestapp.com/oauth2/authorize?client_id=...&redirect_uri=https%3A%2F%2Fscript.google.com%2Fmacros%2Fd%2FPROJECT CODE HERE%2Fusercallback

---

Non-OAuth2 authentication:

HTTP Basic Authentication requires your username and password stored in plain text which is simpler, and its header looks like this:

var auth = Utilities.base64Encode("email@example.com:someharvestpassword");
var headers = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "Authorization": "Basic " + auth
}

If you want Basic authentication, you can ditch startService(), getHarvestService(), authCallback(), and clearService().

You also don't need the OAuth2 library. Just modify the header creation.
