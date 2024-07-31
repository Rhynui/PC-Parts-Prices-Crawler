# PC Parts Prices Crawler

Crawls for prices of specified items on various websites and notifies for any price updates.

## Run

### Test Mode

Run in test mode:

```commandline
python "Shopping Parts Prices Crawler.py"
```

In test mode, the program with run for ever, and prices of each item would be printed to the console continuously.

### Normal Mode

Run in normal mode:

```commandline
python "Shopping Parts Prices Crawler.py" <dir>
```

`dir`: the excel file that the prices would be updated to.

In normal mode, a notification would be printed and the excel file would be modified only when a price change occurs.
