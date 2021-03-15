# ClearSyncIssues-on-Outlook

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/304faa1b4d804c37853095f69139eef0)](https://app.codacy.com/gh/tkoppop/ClearSyncIssues-on-Outlook?utm_source=github.com&utm_medium=referral&utm_content=tkoppop/ClearSyncIssues-on-Outlook&utm_campaign=Badge_Grade_Settings)

When using an online exchange server to manage emails for your custom domain, you will recieve lots of errors such as sync issues. These sync issues are under 4 categories, sync issues, conflicts, local failiures, and server failiures. Sync Issues and conflicts are just duplicates of emails that were unable to be syncronized with the server. Local Failiures are generally emails with attachments >35mb  but <50 mb. This is due to the fact that old outlook emails were able to attach 50mbs of data, while it is not limited to 35 mb. So when a large attachment tries to synchronize it will get rejected. 

This script clears the folders called Sync issues, conflicts, and local failiures. It will permanently delete the emails so they are unrecoverable from recently deleted.
