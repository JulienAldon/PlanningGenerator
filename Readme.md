# Planning Generator for Epitech Intranet

## [WORK IN PROGRESS]

The purpose of the script is to generate plannings accordingly to data on epitech's intranet.

It fetches intranet modules and activities informations then display them in a humain readable format.

For now the format is xlsx only, creating an excel formated planning.

In a future time i'd like to add an ical format support.

# Usage
```
Usage:
        pipenv run python Planning.py <mode> <token> <title> <pedago>
<mode> :
        - ics
        - xlsx
<token> : intranet autologin token
<title> : name of the planning
<pedago> : email address of the filtered pedago
```

# Specifications
- Range of 15 months by default august (08) till november (11) of the next year.

- Calculate the accademic year by default

- List all modules by default 
(`/course/filter?format=json&preload=1&location[]=FR&location[]=FR/LYN&course[]=Code-And-Go&course[]=Dev-And-Go&course[]=bachelor/classic&course[]=premsc&course[]=webacademie&scolaryear[]=<accademic_year>`)
    * filter by :
        - pedago
        - modules
        - promo

- List Projects only manual filter to suppress non project activities (`/module/<accademic_year>/<module['code']>/<module['codeinstance']>/?format=json`)

- Hours specified by default : 
    - TAG ALL DAY

# To add to intranet interface :
- get_activities
- get_modules

# Export to Xlsx planning
- write_cell_merge -> only header
- write_cells -> only header
- write_range -> will add a range of cells colored

# Export to iCal (ics) planning
- write_range -> will add an event to ics calendar