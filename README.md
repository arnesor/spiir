# spiir
Create an Excel sheet overview of spending from https://spiir.com/

## What is Spiir?
Spiir is a free personal finance app that gives you an automatic overview of your
finances and helps you manage your money. It connects to most european bank accounts
and automatically analyses and categorises everything for you.

In addition to the app, Spiir also has a web interface with more advanced
functionality, see https://mine.spiir.dk/. Among other things you can find a good
budgeting function.

## What does this source code do?
If you want more advanced analysis than the app or web-interface provides, this repo is
for you. Based on a csv export from Spiir it creates an Excel sheet with your
spendings, categorised by category and month.

## How to export the necessary csv file from Spiir?
1. Log in to https://mine.spiir.dk/
2. Select your name at the upper right corner and then "Eksporter data".
3. Under the "Avanceret Eksport" heading, click on "Eksporter til CSV".
4. Save the file to the root directory of the cloned spir repo.

## Usage
