# The-Big-Beautiful-Budgeter

Reads bank/cc statements and categorizes them based off of `categories.json` + some excel wizardry to give a nice overview of how we are spending money. Now I don't have to pay a company to do it while also risking they sell my data/get hacked!

put csv from cc into /statements
`python3 main.py` to categorize
run `python3 main.py --rebuild` after making any updates to `categories.json`

Needs:
\*openpyxl

In case I forget, startup:
source .venv/bin/activate
