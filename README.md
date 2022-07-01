# helvetius-cohorts-to-moodle
Export students *xls* listing from Helvetius, preconfigure cohorts listing (a json file), and make *csv* files for upload users by *csv* with cohorts to Moodle.

Place the *xls* file(s) from Helvetius on the same folder of the script and *json* config files.

You can create one *csv* from multiple *xls* files simultaneous.

# Libraries used
This script use the folliwing libraries :
- json
- pandas
- csv
- OleFileIO_PL
- glob

# CSV Moodle Format

| username            | lastname               | firstname | email | cohort1        |auth   |
|---------------------|------------------------|-----------|-------|----------------|-------|
| mail of the student | 1st letter capitalized |           |       | chosen cohort  |oauth2 |
