# kickstarter_scraper
Scrapes the Kickstarter Data into an excel sheet for further analysis.

## Installation
To install the required dependencies, run the command <pre>pip install -Ur requirements.txt</pre>

## Running
To run the script, run the command <pre>python scraper.py</pre>

You will need to enter the category ID of the category you want the data for. It will be a number which can be found at the search query of your kickstarter search.
For Example,
<pre>https://www.kickstarter.com/discover/advanced?category_id=334&sort=magic&seed=2539214&page=1</pre>
The category ID is 334. This ID should be entered in the program when asked for it. Also enter the amount of pages you want to scrape when asked for it.<br>
The program will execute once you enter the Category ID and choose the number of pages to scrape and a new excel file will be created each time the script is run.
