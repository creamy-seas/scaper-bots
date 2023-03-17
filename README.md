# Outlook bot #

**Logging onto an outlook account and scrapping up email in the inbox**

-------------------------------------------------------------------------------

- `pandas_out` is the Dataframe holding scraped content in columns *From* *Subject* *Date* *Content* *Forwarded Content*

- Output is written to the <kbd>output</kbd> folder in either `csv` or `pkl` format

- Can select emails 

    > By read/unread
	
	> Date
	
	> Position in inbox
	
- By default, the whole directoy is scraped, unless `outlook_scrape_filters(...)` is run to set

- **The browser window must be visible on screen at all times**


-------------------------------------------------------------------------------

## Calling procedure ##

- With specified dates:
``` python
########################################
########################################
outlook_id="outlook@sbtgc.local"
password="1488babab"
timeout=50
browser="chrome"
date_min = [1, 3, 2019]
date_max = [2, 3, 2020]
########################################
########################################
running_class = outlook_bot("chrome", timeout)
running_class.outlook_login(outlook_id, password)
running_class.outlook_scrape_filters(date_min, date_max, scrape_only_unread=True)
running_class.outlook_scrape("data", "csv")
```

- Email in the inbox from 0 to 20 inclusive
``` python
########################################
########################################
outlook_id="outlook@sbtgc.local"
password="1488babab"
timeout=50
browser="chrome"
date_min = None
date_max = None
########################################
########################################
running_class = outlook_bot("chrome", timeout)
running_class.outlook_login(outlook_id, password)
running_class.outlook_scrape_filters(date_min, date_max, scrape_only_unread=True, 0, 20)
running_class.outlook_scrape("data", "csv")
```
