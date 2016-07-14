# O365-CustomAction-NodeJS
Build SharePoint Custom Item Context Menu Action using nodeJs
##Make this sample work
1. Create apponly app principle using /_layouts/15/appregnew.aspx
2. Update O365Constants.js with client_id and client_secret for app_only app to get a pair navigate to https://{yourdomain}.sharepoint.com/_layouts/15/AppRegNew.aspx 
make sure that the generated credentials will have update list item permission
4. replace your actual subdomain in O365Constants.js
5. run <br/><code>npm install</code><br/>
6. create azure website or use ngrok to test using your local machine
7. link azure website to your forked github repo 
8. push your new changes to your repository , gulp task deploy-sharepoint will be triggered as part of azure deployment script deploy.cmd

Try it out it works on Custom lists only and it updates the OTB title to "Title updated by NodeJS Custom Action"
![example](./example.png)