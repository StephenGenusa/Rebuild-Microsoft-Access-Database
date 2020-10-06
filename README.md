# Rebuild Microsoft Access Database

Rebuild Microsoft Access DB with a single command. Handles forms, reports, modules, and queries.

Some day, Microsoft Access will be a distant memory. In the meantime we have "Compact and Repair" and the need to rebuild a database from exported text files -- a *decades* old problem enhanced by multi-thread corruption.

rebuildaccessdb.py 


- -i --input-file
- -d --download-script

rebuildaccessdb.py -i msdpt.accdb

This command will copy the db to a work directory, delete all forms, reports, modules and queries, leaving the table definitions. After C&R, exports are made from the original file and then imported into the new file. After import, more C&R to shrink the file.

Uses the [https://github.com/vbaidiot/Ariawase](https://github.com/vbaidiot/Ariawase) vbac.wsf script to do some of the lifting.

