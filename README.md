# Rotamap Sessions to Outlook Calendar
Uses Rotamap API to populate an Outlook calendar with clinical sessions. Works for Medirota or CLWRota. 

## Description
This uses the Rotamap API to get clinical sessions for a given clinician and populate their Outlook calendar.  This ensures that their Free/Busy time in Outlook is accurate and makes appointment planning simpler for the EA/PA as all appointments and clinical commitments can be found in one calendar. Details of Rotamap Public API: https://www.rotamap.net/docs/publicapi
On subsequent runs, all previously-created sessions on future dates are deleted and then rewritten.

## Getting Started

### API Setup
The Super user account on CLWRota/Medirota is used to configure API user accounts.  Once logged into the Super User account, go to User Admin > API users to configure an API account. This is required to generate a token and subsequent API call.

### Variables

Most of my variable information is stored in a text file in the Python directory - this isn't ideal but I don't have administrator access to modify environment variables on my work PC which would be preferable. You can see in the code that the only environment variables set are the API username and password. 
The format of the text document is as follows:
<br>[Clinician Name]
<br>[Clinician Email Address]
<br>[Site URL inlcuding /publicapi] e.g. https://ex.clwrota.com/publicapi
<br>[API Login URL] e.g. https://ex.clwrota.com/publicapi
<br>[Proxies] _required by my organisation_

### Executing program

Best run using Task Scheduler or similar outside of work hours.

## Authors

Courtney Russ
<br>https://www.linkedin.com/in/courtney-russ/
<br>crus047@aucklanduni.ac.nz / cr285@students.waikato.ac.nz

## Version History
* 0.1
    * Initial Release

## License

This project is licensed under the GNU General Public License v3.0 - see the LICENSE.md file for details
