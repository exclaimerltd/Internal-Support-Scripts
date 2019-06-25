### Usage of the Bulk User Create Script

---

#### Scenario

You have configured a lab environment and require a large number of users to test the Exclaimer software correctly without replicating your production environment.

#### Steps to use script

1. Copy both files to a folder on your machine
2. Open command prompt as an administrator
3. Run *cscript adgen.vbs* in the command prompt window
4. The script will prompt for the following
 a. Company Name - This will create a top level OU
 b. How many users do you require - To a maximum of 500k objects
 c. How many Countries do your require - This will create OUs within the Company OU
 d. How many Offices do you need - This will create under each Country OU
 e. How many Departments are required - Created within the Office OUs
5. It will finally ask if you want a log file created containing the users

Once completed, a refresh of the AD information will show the data randomly created by the application.  OUs for country, office and department will be created also.