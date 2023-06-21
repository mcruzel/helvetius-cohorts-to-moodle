# helvetius-cohorts-to-moodle
Export students *xls* listing from Helvetius, preconfigure cohorts listing (a json file), aud automatically create accounts on Moodle and put them in the aimed cohorts. **You need to create cohorts before**, the script will search cohorts by name (**by default : use system-wide cohorts**).

Place the *xls* file(s) from Helvetius on the same folder of the script and *json* config files.

You use this script with multiple *xls* files simultaneous.

# Prerequisite
This script use a sendgrid API key for sending mail when an account is created. If you don't use Sengrid, you need to disable *mail_function()* function call

You'll probably need to edit the mail content on *mail_function()*

Edit the *json* config file with Helvetius objects :
- Formation : the "global" curriculum (example : the Bachelor's master's doctorate system system in European Union)
- Site : the location
- Produit : the curriculum (example : 1st year of Mathematics)
- Moodle cohort name : the exact cohortname (on Moodle) where you want to put students, corresponding to their group
- Moodle cohort merge to : same but with a global curriculum group

  Create a Moodle account ((https://docs.moodle.org/400/en/Using_web_services)) and allow him to use the following webservices :
  - core_user_get_users
  - core_user_create_users
  - core_cohort_search_cohorts
  - core_cohort_add_cohort_members

# Libraries used
This script use the folliwing libraries :
- requests
- time
- json
- urllib.parse
- olefile
- sendgrid
- os
- sendgrid.helpers.mail
- glob
- pandas

# Helvetius files

*xls* files are obtained from Helvetius :
Inscription > Extraire les étudiants

Export format : Excel.
